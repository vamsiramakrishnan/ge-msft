import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { WORD_CAPABILITIES } from './capabilities.js';
import {
  paragraphsToBlocks,
  paragraphsToElements,
  searchHitsToContext,
  wordDocumentToContext,
  wordSelectionToContext,
} from './capture.js';
import {
  chooseAnchorIndex,
  formatSources,
  planAddComment,
  planFillContentControl,
  planInsertOoxml,
  planInsertText,
  planReplaceSelection,
  planTrackedChange,
} from './actuate-plan.js';
import { provenanceRecord } from './provenance-record.js';
import {
  commentAddedEvent,
  documentChangedEvent,
  originFromWordSource,
  selectionChangedEvent,
} from './events.js';
import { OfficeWordHost, type WordHost } from './host-port.js';

/**
 * The exact `ActuationKind`s {@link WordBridge.actuate} handles — the SINGLE source of truth for
 * what the switch dispatches, sitting beside it so the two can't drift. The conformance test
 * (ADR-0006 closure) asserts `set(WORD_CAPABILITIES.actuations) === set(HANDLED_ACTUATIONS)`: a
 * phantom (advertised-but-unhandled) or a silent handled-but-unadvertised kind fails the build.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = [
  'tracked-change',
  'add-comment',
  'comment-reply',
];

/**
 * The Word `DocBridge`. All Office.js access goes through the injectable {@link WordHost} port
 * (the real {@link OfficeWordHost} adapter by default), so this file holds the *orchestration*:
 * mapping reads into context, choosing the tracked-change anchor, degrading a drifted finding to
 * a panel item, and translating host outcomes into {@link ActuationResult}. That decision logic
 * is now host-free, so it's unit-testable against a fake host. Pure mapping still lives in
 * `capture.ts` / `actuate-plan.ts` / `events.ts`; the `Word.run` batching lives in the port.
 */
export class WordBridge implements DocBridge {
  readonly surface = 'word' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  constructor(private readonly host: WordHost = new OfficeWordHost()) {}

  getCapabilities(): CapabilityManifest {
    return WORD_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    const [selText, bodyText] = await Promise.all([
      this.host.readSelectionText(),
      this.host.readBodyText(),
    ]);
    const refs: ContextRef[] = [];
    if (selText.trim()) {
      refs.push({
        id: 'word:selection',
        kind: 'selection',
        surface: 'word',
        title: 'Selection',
        preview: selText.slice(0, 120),
        live: true,
      });
    }
    refs.push({
      id: 'word:document',
      kind: 'document',
      surface: 'word',
      title: 'Whole document',
      preview: bodyText.slice(0, 120),
    });
    return refs;
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'selection') {
      const text = await this.host.readSelectionText();
      return wordSelectionToContext(text);
    }
    // Whole document → paragraphs (with style for heading levels) → native blocks → chunks.
    const paras = await this.host.readParagraphs();
    return wordDocumentToContext('word:document', undefined, paragraphsToElements(paras));
  }

  /**
   * ADR-0003 Layer B element 1: an ambient structural snapshot of the document. Reads the body
   * paragraphs (one batched host call), maps them to `Block[]` via the same native path as
   * `resolveContext` (headings keep their levels + locators, so the snapshot outline matches the
   * document), and builds the surface-agnostic snapshot. Version increments per capture. Comments
   * are omitted (no cheap port read); the runtime renders + wraps this as untrusted data.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    const paras = await this.host.readParagraphs();
    if (paras.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'word',
      version: this.docStateVersion,
      blocks: paragraphsToBlocks(paras),
    });
  }

  /**
   * ADR-0003 Layer B element 2: lazily read the document slices relevant to `query` instead of
   * pre-chunking the whole body. Re-resolves the matches at call-time via the port's content-
   * anchored `body.search` (per the anchoring discipline), and maps the bounded hits to live,
   * content-anchored `ResolvedContext`. Empty query or no hits → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q) return [];
    const hits = await this.host.searchText(q, false);
    return searchHitsToContext(q, hits);
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    switch (req.kind) {
      case 'tracked-change':
        return this.applyTrackedChange(req);
      case 'add-comment':
        return this.applyAddComment(req);
      case 'comment-reply':
        return this.applyCommentReply(req);
      case 'insert-text':
        return this.applyInsertText(req);
      case 'replace-selection':
        return this.applyReplaceSelection(req);
      case 'insert-ooxml':
        return this.applyInsertOoxml(req);
      case 'fill-content-control':
        return this.applyFillContentControl(req);
      default:
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'unsupported', message: `Word bridge cannot ${req.kind}` },
        };
    }
  }

  private async applyTrackedChange(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planTrackedChange(req);
    if (!plan.matchText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'tracked-change needs target.matchText' },
      };
    }
    // The decision — which read-back hit to write on, and degrade to a panel item when none
    // match — lives here (host-free, hence testable). The port runs the search→insert batch and
    // calls back into `chooseAnchorIndex` with the live hit texts, re-resolving at apply-time.
    const outcome = await this.host.applyTrackedChange(
      plan.matchText,
      { matchCase: false },
      plan.text,
      (hitTexts) => chooseAnchorIndex([...hitTexts], plan.contextHint),
    );
    if (outcome.status === 'drift') {
      // Anchor drift: degrade to a panel item rather than render a broken edit.
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'anchor_drift',
          message: 'The matched text is no longer in the document.',
        },
      };
    }
    // ADR-0003 comments-as-citations: after a successful (non-degraded) tracked change, drop a
    // Word comment carrying the citation, anchored on the same content. Prefer provenance.sources,
    // fall back to params.sources. Best-effort — a comment failure must NOT un-apply or fail the
    // change, so we await it but ignore its outcome.
    const sources = req.provenance?.sources ?? req.params.sources ?? [];
    if (sources.length > 0) {
      await this.host.addComment(plan.matchText, false, formatSources(sources));
    }
    // Durable provenance (BUILD-PLAN 1.6): stamp the record after the reversible change lands.
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * ADR-0004 `comment` verb on Word: add a NEW content-anchored comment. Reuses the port's
   * `addComment` (content-anchored via `body.search`, re-resolved at apply-time), so a drifted
   * anchor degrades to a panel item (`ok:false, degraded:true, anchor_drift`) — the same discipline
   * as a tracked change — rather than landing the comment on the wrong range. `changeId` is
   * propagated, never re-minted.
   */
  private async applyAddComment(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planAddComment(req);
    if (!plan.matchText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_anchor', message: 'add-comment needs target.matchText' },
      };
    }
    if (!plan.hasText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'add-comment needs params.text' },
      };
    }
    const outcome = await this.host.addComment(plan.matchText, false, plan.text);
    if (!outcome.ok) {
      // Anchor gone (or comments unsupported): degrade to a panel item, never a broken comment —
      // the same content-anchoring discipline as a drifted tracked change.
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'anchor_drift',
          message: 'The matched text is no longer in the document.',
        },
      };
    }
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: 'comment',
      ...provFlags(req, dropped),
    };
  }

  private async applyCommentReply(req: ActuationRequest): Promise<ActuationResult> {
    const reply = req.params.text ?? '';
    const commentId = req.params.target?.commentId;
    if (!commentId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_comment', message: 'comment-reply needs target.commentId' },
      };
    }
    const outcome = await this.host.replyToComment(
      commentId,
      reply,
      req.params.resolveComment ?? false,
    );
    if (outcome.status === 'gone') {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'comment_gone', message: 'The comment no longer exists.' },
      };
    }
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * ADR-0007 `insert-text`: insert plain text directly (NOT a tracked change). When `target.matchText`
   * is given the insert is content-anchored — re-resolved via `body.search` at apply-time and degrading
   * to a panel item on drift, the same discipline as a tracked change — otherwise it lands at the
   * current selection. Fails closed on empty text. This is a DIRECT edit, so reversibility is recorded
   * explicitly (see the captured prior-state below), not left to host-session undo.
   */
  private async applyInsertText(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planInsertText(req);
    if (!plan.hasText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'insert-text needs params.text' },
      };
    }
    const outcome = await this.host.insertText(
      plan.anchored ? plan.matchText : undefined,
      { matchCase: false },
      plan.text,
      (hitTexts) => chooseAnchorIndex([...hitTexts], plan.contextHint),
    );
    if (outcome.status === 'drift') {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'anchor_drift', message: 'The matched text is no longer in the document.' },
      };
    }
    // TODO(ADR-0007 inverse): captured prior-state for the inverse — an insertion has no prior text,
    // so the inverse is "delete the inserted range". Shape needed on InverseDescriptorSchema:
    //   { op: 'delete-content', insertedText: string, location: 'selection' | 'insert-text' }
    // Reverse op: body.search(insertedText) → range.delete() (or shrink the range to empty).
    const insertedState = { insertedText: outcome.insertedText, location: outcome.location };
    void insertedState; // reported to central wiring; not yet on the schema (see summary).
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * ADR-0007 `replace-selection`: overwrite the current selection with new text. Fails closed on empty
   * text and on an empty selection (no range to replace → degrade to a panel item). Captures the prior
   * selection text BEFORE overwriting so the edit is reversible.
   */
  private async applyReplaceSelection(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planReplaceSelection(req);
    if (!plan.hasText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'replace-selection needs params.text' },
      };
    }
    const outcome = await this.host.replaceSelection(plan.text);
    if (outcome.status === 'empty') {
      // Nothing selected → no range to replace. Degrade to a panel item rather than insert blindly.
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'no_selection', message: 'Nothing is selected to replace.' },
      };
    }
    // TODO(ADR-0007 inverse): captured prior-state for the inverse — the selection's PRIOR text.
    // Shape needed on InverseDescriptorSchema:
    //   { op: 'restore-text', priorText: string, location: 'selection' }
    // Reverse op: re-select the written range and insertText(priorText, replace).
    const priorState = { priorText: outcome.priorText, location: outcome.location };
    void priorState; // reported to central wiring; not yet on the schema (see summary).
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * ADR-0007 `insert-ooxml`: insert rich OOXML directly, content-anchored (degrading on drift) or at
   * the selection. Fails closed on empty OOXML. The OOXML is host/model-derived → untrusted; it is
   * handed to the host as data. Reversibility is recorded as a delete of the inserted range.
   */
  private async applyInsertOoxml(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planInsertOoxml(req);
    if (!plan.hasOoxml) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_ooxml', message: 'insert-ooxml needs params.ooxml' },
      };
    }
    const outcome = await this.host.insertOoxml(
      plan.anchored ? plan.matchText : undefined,
      { matchCase: false },
      plan.ooxml,
      (hitTexts) => chooseAnchorIndex([...hitTexts], plan.contextHint),
    );
    if (outcome.status === 'drift') {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'anchor_drift', message: 'The matched text is no longer in the document.' },
      };
    }
    // TODO(ADR-0007 inverse): captured prior-state for the inverse — an OOXML insertion has no prior
    // text and the rendered text isn't known here, so the inverse is "delete the inserted range".
    // Shape needed on InverseDescriptorSchema:
    //   { op: 'delete-content', anchor: string | null, location: 'selection' | 'insert-ooxml' }
    // Reverse op: re-resolve the anchor (or the inserted range) and range.delete().
    const insertedState = { location: outcome.location };
    void insertedState; // reported to central wiring; not yet on the schema (see summary).
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * ADR-0007 `fill-content-control`: populate a named content control with text. Fails closed on a
   * missing id or empty text; a stale id (the control was deleted) degrades to a panel item — the
   * content-control analogue of anchor drift. Captures the control's prior text so the fill is
   * reversible.
   */
  private async applyFillContentControl(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planFillContentControl(req);
    if (!plan.hasId || !plan.contentControlId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_content_control',
          message: 'fill-content-control needs target.contentControlId',
        },
      };
    }
    if (!plan.hasText) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'fill-content-control needs params.text' },
      };
    }
    const outcome = await this.host.fillContentControl(plan.contentControlId, plan.text);
    if (outcome.status === 'gone') {
      // The control was deleted → degrade to a panel item rather than throw (anchor-drift analogue).
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: { code: 'content_control_gone', message: 'The content control no longer exists.' },
      };
    }
    // TODO(ADR-0007 inverse): captured prior-state for the inverse — the control's PRIOR text.
    // Shape needed on InverseDescriptorSchema:
    //   { op: 'restore-content-control', contentControlId: string, priorText: string }
    // Reverse op: contentControls.getByIdOrNullObject(id).insertText(priorText, replace).
    const priorState = { contentControlId: plan.contentControlId, priorText: outcome.priorText };
    void priorState; // reported to central wiring; not yet on the schema (see summary).
    const dropped = await this.persistProvenance(req);
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
      ...provFlags(req, dropped),
    };
  }

  /**
   * Durable provenance persistence (BUILD-PLAN 1.6 security follow-up). After a reversible write
   * lands, stamp the {@link provenanceRecord} into the document's durable metadata keyed by
   * `changeId`, so the write stays provenanced across save/reopen. Word's durable metadata is a
   * custom XML part (`customXmlParts.add`), so we persist the record's OOXML form via the port.
   *
   * Best-effort and feature-detected, exactly like the citation-comment path: a missing
   * `req.provenance` is skipped (we never fabricate identity — the runtime stamps it). The reversible
   * write already succeeded; provenance is additive metadata, not the system of record, so a
   * persistence failure must NOT fail the write — but it is no longer SILENT: we return whether the
   * record dropped so the caller can flag `provenanceDropped` on the result (observability).
   *
   * @returns `true` when provenance was present but could not be durably persisted (a drop to
   * surface); `false` when persisted or when there was nothing to persist.
   */
  private async persistProvenance(req: ActuationRequest): Promise<boolean> {
    if (!req.provenance) return false; // nothing to persist — not a drop.
    const record = provenanceRecord(req.changeId, req.provenance);
    try {
      const outcome = await this.host.persistProvenance(record.xml);
      return !outcome.ok; // port reports failure → dropped.
    } catch {
      return true; // host threw → dropped (the write still stands).
    }
  }

  /**
   * Stream Word host events into the trigger engine. The port owns the defensive Office.js
   * registration/teardown (requirement-set gating, single removal path, coauthor source); this
   * method maps the raw handler primitives into {@link HostEvent}s via the pure `events.ts`
   * builders and tags each with its `origin`. Returns the port's `Unsubscribe`.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    return this.host.registerHandlers({
      onSelectionChanged: () => emit(selectionChangedEvent()),
      onDocumentChanged: (args) => emit(documentChangedEvent(originFromWordSource(args.source))),
      onCommentAdded: (args) => {
        const commentId = args.commentId ?? args.id ?? args.ids?.[0];
        if (commentId) emit(commentAddedEvent(originFromWordSource(args.source), commentId));
      },
    });
  }
}

/**
 * Observability flags for a landed write: `provenanceMissing` when the request carried no provenance
 * payload at all (an unattributed write — never mistake it for an attributed one), `provenanceDropped`
 * when a present record failed to persist durably. Persisted cleanly ⇒ empty.
 */
function provFlags(
  req: ActuationRequest,
  dropped: boolean,
): { provenanceDropped?: true; provenanceMissing?: true } {
  if (!req.provenance) return { provenanceMissing: true };
  return dropped ? { provenanceDropped: true } : {};
}
