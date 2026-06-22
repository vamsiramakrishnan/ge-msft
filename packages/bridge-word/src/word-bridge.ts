import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { WORD_CAPABILITIES } from './capabilities.js';
import {
  headingLevel,
  wordDocumentToContext,
  wordSelectionToContext,
  type WordElement,
} from './capture.js';
import { chooseAnchorIndex, planTrackedChange } from './actuate-plan.js';
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
    const elements: WordElement[] = paras.map((p) => {
      const level = headingLevel(p.styleBuiltIn);
      return level > 0
        ? { kind: 'heading' as const, text: p.text, level }
        : { kind: 'paragraph' as const, text: p.text };
    });
    return wordDocumentToContext('word:document', undefined, elements);
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
