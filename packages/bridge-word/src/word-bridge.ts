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

/**
 * The Word `DocBridge`. The ONLY place Office.js (`Word.run`) is touched. Reads via the
 * native object model (selection, paragraphs, styles), writes via **tracked changes anchored
 * by content** (`body.search`, re-resolved at apply-time → degrade if drifted). Pure mapping
 * lives in `capture.ts` / `actuate-plan.ts` (unit-tested); this file is the host wiring.
 */
export class WordBridge implements DocBridge {
  readonly surface = 'word' as const;

  getCapabilities(): CapabilityManifest {
    return WORD_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      const refs: ContextRef[] = [];
      if (sel.text.trim()) {
        refs.push({
          id: 'word:selection',
          kind: 'selection',
          surface: 'word',
          title: 'Selection',
          preview: sel.text.slice(0, 120),
          live: true,
        });
      }
      refs.push({
        id: 'word:document',
        kind: 'document',
        surface: 'word',
        title: 'Whole document',
        preview: body.text.slice(0, 120),
      });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'selection') {
      return Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load('text');
        await ctx.sync();
        return wordSelectionToContext(sel.text);
      });
    }
    // Whole document → paragraphs (with style for heading levels) → native blocks → chunks.
    return Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items/text,items/styleBuiltIn');
      await ctx.sync();
      const elements: WordElement[] = paras.items
        .filter((p) => p.text.trim().length > 0)
        .map((p) => {
          const level = headingLevel(String(p.styleBuiltIn));
          return level > 0
            ? { kind: 'heading' as const, text: p.text, level }
            : { kind: 'paragraph' as const, text: p.text };
        });
      return wordDocumentToContext('word:document', undefined, elements);
    });
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
    return Word.run(async (ctx) => {
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      const results = ctx.document.body.search(plan.matchText!, { matchCase: false });
      results.load('items/text');
      await ctx.sync();

      const idx = chooseAnchorIndex(
        results.items.map((r) => r.text),
        plan.contextHint,
      );
      const range = idx >= 0 ? results.items[idx] : undefined;
      if (!range) {
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
      range.insertText(plan.text, Word.InsertLocation.replace);
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: 'tracked-change' };
    });
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
    return Word.run(async (ctx) => {
      const comments = ctx.document.body.getComments();
      comments.load('items/id');
      await ctx.sync();
      const comment = comments.items.find((c) => c.id === commentId);
      if (!comment) {
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          degraded: true,
          error: { code: 'comment_gone', message: 'The comment no longer exists.' },
        };
      }
      comment.reply(reply);
      if (req.params.resolveComment) comment.resolved = true;
      await ctx.sync();
      return { ok: true, changeId: req.changeId, kind: req.kind, location: `comment:${commentId}` };
    });
  }

  /**
   * Stream Word host events into the trigger engine. Each registration is defensive: not every
   * event exists on every Word requirement set, so a failed/absent registration simply means we
   * never emit that event — it never throws. Coauthor (remote) edits are tagged so the registry
   * drops them by default. Returns an `Unsubscribe` that removes *every* handler we added.
   *
   * Confirmed against node_modules/@types/office-js/index.d.ts:
   *   - Office selection: `Office.EventType.DocumentSelectionChanged` (l.645) +
   *     add/removeHandlerAsync (l.3875 / l.3965). No coauthor source → origin always 'local'.
   *   - Word edits: `Document.onParagraphChanged` (l.102848) / `onParagraphAdded` (l.102839) /
   *     `onParagraphDeleted` (l.102857); args carry `source: Word.EventSource` (l.118797 etc.).
   *   - `Word.EventSource` enum 'Local' | 'Remote' (l.118481). Word `EventHandlers.add` returns an
   *     `EventHandlerResult` whose `.remove()` (l.25582) must run inside a `Word.run` + sync.
   *   - The Word `Document` type in this typings version has NO `onCommentAdded`, so comments are
   *     feature-detected at runtime and skipped when absent.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    const teardowns: Array<() => void> = [];

    // --- Selection (Office host event; no coauthor source → local) ---
    try {
      const onSelection = (): void => {
        emit(selectionChangedEvent());
      };
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelection,
      );
      teardowns.push(() => {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
            handler: onSelection,
          });
        } catch {
          // best-effort: removal may fail if the host already tore the handler down.
        }
      });
    } catch {
      // Selection observation unavailable on this host — simply don't emit it.
    }

    // --- Document edits + comments (Word object-model events; carry coauthor source) ---
    // Registration runs async inside Word.run; the returned EventHandlerResults are collected so
    // the unsubscribe can remove them (also inside a Word.run, per the typings' note).
    const removers: Array<{ remove(): void }> = [];
    let unsubscribed = false;

    void Word.run(async (ctx) => {
      const doc = ctx.document;

      const onDocChange = (args: {
        source: Word.EventSource | 'Local' | 'Remote';
      }): Promise<void> => {
        emit(documentChangedEvent(originFromWordSource(args.source)));
        return Promise.resolve();
      };
      for (const handlers of [
        doc.onParagraphChanged,
        doc.onParagraphAdded,
        doc.onParagraphDeleted,
      ]) {
        try {
          removers.push(handlers.add(onDocChange));
        } catch {
          // This paragraph event isn't in the active requirement set — skip it.
        }
      }

      // Comments: not present in every Word typings/requirement set. Feature-detect off `unknown`.
      const maybeComment = (doc as unknown as Record<string, unknown>).onCommentAdded;
      if (isCommentHandlers(maybeComment)) {
        const onComment = (args: {
          source?: Word.EventSource | 'Local' | 'Remote';
          commentId?: string;
          id?: string;
          ids?: string[];
        }): Promise<void> => {
          const commentId = args.commentId ?? args.id ?? args.ids?.[0];
          if (commentId) {
            emit(commentAddedEvent(originFromWordSource(args.source), commentId));
          }
          return Promise.resolve();
        };
        try {
          removers.push(maybeComment.add(onComment));
        } catch {
          // Comment events unavailable — skip.
        }
      }

      await ctx.sync();

      // If unsubscribe already fired before registration completed, tear down immediately.
      if (unsubscribed) await this.removeWordHandlers(removers.splice(0));
    }).catch(() => {
      // Word.run failed (no host / unsupported) — nothing registered, nothing to clean up.
    });

    teardowns.push(() => {
      unsubscribed = true;
      const toRemove = removers.splice(0);
      if (toRemove.length > 0) void this.removeWordHandlers(toRemove);
    });

    return () => {
      for (const teardown of teardowns) teardown();
    };
  }

  /** Remove Word object-model event handlers; per the typings, `.remove()` runs inside a sync batch. */
  private async removeWordHandlers(handlers: Array<{ remove(): void }>): Promise<void> {
    try {
      await Word.run(async (ctx) => {
        for (const h of handlers) {
          try {
            h.remove();
          } catch {
            // individual handler may already be gone
          }
        }
        await ctx.sync();
      });
    } catch {
      // best-effort teardown
    }
  }
}

/** Narrow an `unknown` document member to something with an `add()` we can register on. */
function isCommentHandlers(
  value: unknown,
): value is { add(handler: (args: never) => Promise<void>): { remove(): void } } {
  return (
    typeof value === 'object' &&
    value !== null &&
    'add' in value &&
    typeof (value as { add: unknown }).add === 'function'
  );
}
