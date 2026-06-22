import type {
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
import { chooseAnchorIndex, formatSources, planTrackedChange } from './actuate-plan.js';
import {
  commentAddedEvent,
  documentChangedEvent,
  originFromWordSource,
  selectionChangedEvent,
} from './events.js';
import { OfficeWordHost, type WordHost } from './host-port.js';

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
      case 'comment-reply':
        return this.applyCommentReply(req);
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
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: outcome.location,
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
    return { ok: true, changeId: req.changeId, kind: req.kind, location: outcome.location };
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
