import { isSet } from './capabilities-runtime.js';

/**
 * The narrow **host port** the {@link WordBridge} drives. It captures ONLY the Word host
 * operations the bridge actually performs — read selection / body / paragraphs, apply a
 * content-anchored tracked change, register host-event handlers — and is deliberately NOT a
 * 1:1 mirror of Office.js. The real {@link OfficeWordHost} translates these calls into the
 * existing `Word.run` batches (preserving the sync-batching and the apply-time anchor
 * re-resolve); a fake implements the same interface so the bridge's orchestration is
 * unit-testable without Office.js.
 *
 * The load-bearing decision (which search hit to pick, when to degrade a drifted finding to a
 * panel item) lives on the *bridge* side. The port exposes that as a single
 * {@link WordHost.applyTrackedChange} call that hands the read-back hit texts to a `choose`
 * callback and inserts on the chosen hit inside the same batch — so the search→insert
 * read-then-write dependency is preserved while the choice stays testable in the bridge.
 */

/** A paragraph read from the document body (text + the built-in style name for heading level). */
export interface WordParagraph {
  readonly text: string;
  readonly styleBuiltIn: string;
}

/** The outcome of an apply-time anchor resolution + insert. */
export type TrackedChangeOutcome =
  | { readonly status: 'applied'; readonly location: string }
  | { readonly status: 'drift' }; // 0 hits, or the chooser rejected every hit

/**
 * Pick which of the read-back search hits to write on. Returns the chosen index, or a negative
 * number to signal anchor drift (degrade to a panel item rather than edit the wrong range).
 * Runs *inside* the host's search→insert batch, after the hits are read back.
 */
export type ChooseHit = (hitTexts: readonly string[]) => number;

/** A reply target found while replying to a comment. */
export type CommentReplyOutcome =
  | { readonly status: 'replied'; readonly location: string }
  | { readonly status: 'gone' };

/** Raw paragraph-edit handler args at the host boundary (coauthor source is untrusted). */
export interface WordEditArgs {
  readonly source: unknown;
}

/** Raw comment-added handler args at the host boundary. */
export interface WordCommentArgs {
  readonly source?: unknown;
  readonly commentId?: string;
  readonly id?: string;
  readonly ids?: readonly string[];
}

/** Callbacks the bridge supplies to {@link WordHost.registerHandlers}; each may be unsupported. */
export interface WordHandlers {
  readonly onSelectionChanged: () => void;
  readonly onDocumentChanged: (args: WordEditArgs) => void;
  readonly onCommentAdded: (args: WordCommentArgs) => void;
}

/**
 * The narrow Word host port. Every method is behavior-focused; the implementations own the
 * `Word.run`/`Office.context` batching, the bridge owns the decisions.
 */
export interface WordHost {
  /** The current selection's text (empty string when nothing is selected). */
  readSelectionText(): Promise<string>;

  /** The whole body text (for the document context preview). */
  readBodyText(): Promise<string>;

  /** Non-empty body paragraphs with their built-in style (for heading-level derivation). */
  readParagraphs(): Promise<WordParagraph[]>;

  /**
   * Turn on tracked changes, search the body for `query`, let `choose` pick a read-back hit, and
   * insert `text` (replace) on the chosen hit — all in one batch. Returns `drift` when there are
   * no hits or `choose` rejects them all (no write happens in that case).
   */
  applyTrackedChange(
    query: string,
    opts: { readonly matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<TrackedChangeOutcome>;

  /** Reply to the comment with `commentId`; optionally resolve it. `gone` if it no longer exists. */
  replyToComment(commentId: string, reply: string, resolve: boolean): Promise<CommentReplyOutcome>;

  /**
   * Register the host-event handlers (selection / paragraph edits / comments). Returns an
   * unsubscribe that removes every handler that attached. Never throws.
   */
  registerHandlers(handlers: WordHandlers): () => void;
}

/**
 * The real adapter: a thin wrapper translating {@link WordHost} calls into the existing
 * `Word.run` / `Office.context` batches. This is the un-unit-tested seam; it preserves the
 * prior bridge's semantics exactly — the read-then-write sync batching, the apply-time anchor
 * re-resolve that degrades a drifted finding, and the defensive event registration/teardown.
 */
export class OfficeWordHost implements WordHost {
  async readSelectionText(): Promise<string> {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      return sel.text;
    });
  }

  async readBodyText(): Promise<string> {
    return Word.run(async (ctx) => {
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return body.text;
    });
  }

  async readParagraphs(): Promise<WordParagraph[]> {
    return Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items/text,items/styleBuiltIn');
      await ctx.sync();
      return paras.items
        .filter((p) => p.text.trim().length > 0)
        .map((p) => ({ text: p.text, styleBuiltIn: String(p.styleBuiltIn) }));
    });
  }

  async applyTrackedChange(
    query: string,
    opts: { readonly matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<TrackedChangeOutcome> {
    return Word.run(async (ctx) => {
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      // First sync is required: the anchor index is chosen from the *read-back* match texts
      // before the second (write) sync inserts on the chosen range — a read-then-write
      // dependency, and the load-bearing re-resolve that degrades a drifted finding to a panel
      // item rather than editing the wrong range.
      await ctx.sync();

      const idx = choose(results.items.map((r) => r.text));
      const range = idx >= 0 ? results.items[idx] : undefined;
      if (!range) return { status: 'drift' };

      range.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      return { status: 'applied', location: 'tracked-change' };
    });
  }

  async replyToComment(
    commentId: string,
    reply: string,
    resolve: boolean,
  ): Promise<CommentReplyOutcome> {
    return Word.run(async (ctx) => {
      const comments = ctx.document.body.getComments();
      comments.load('items/id');
      await ctx.sync();
      const comment = comments.items.find((c) => c.id === commentId);
      if (!comment) return { status: 'gone' };
      comment.reply(reply);
      if (resolve) comment.resolved = true;
      await ctx.sync();
      return { status: 'replied', location: `comment:${commentId}` };
    });
  }

  registerHandlers(handlers: WordHandlers): () => void {
    // --- Selection (Office host event; no coauthor source → local) ---
    // Registered synchronously; its removal is folded into the single settled teardown path below
    // so there is exactly one owner of teardown (no second, racing removal path).
    let onSelection: (() => void) | undefined;
    try {
      const handler = (): void => handlers.onSelectionChanged();
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
      onSelection = handler;
    } catch {
      // Selection observation unavailable on this host — simply don't emit it.
    }

    const removeSelection = (): void => {
      if (!onSelection) return;
      try {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
          handler: onSelection,
        });
      } catch {
        // best-effort: removal may fail if the host already tore the handler down.
      } finally {
        onSelection = undefined;
      }
    };

    // --- Document edits + comments (Word object-model events; carry coauthor source) ---
    // Registration runs async inside Word.run; the returned EventHandlerResults are collected so
    // the unsubscribe can remove them (also inside a Word.run, per the typings' note).
    const removers: Array<{ remove(): void }> = [];
    let unsubscribed = false;

    // Mirror the Excel bridge: capture the registration promise and have the teardown chain off
    // it. The single removal path lives in that settled `.then`, so handlers committed on the host
    // after a synchronous unsubscribe are still removed exactly once — no leak, no double-splice.
    const registration = Word.run(async (ctx) => {
      const doc = ctx.document;

      const onDocChange = (args: {
        source: Word.EventSource | 'Local' | 'Remote';
      }): Promise<void> => {
        handlers.onDocumentChanged({ source: args.source });
        return Promise.resolve();
      };
      // PRIMARY gate is the requirement-set check, NOT property truthiness: `onParagraph*` are
      // always-truthy getters on the Office.js proxy, so a truthiness check gates nothing and
      // `.add()` THROWS on a host below the supporting set. All three paragraph events share one
      // requirement set: `Document.onParagraphAdded/Changed/Deleted` → WordApi 1.6 (typings
      // l.102835 / l.102844 / l.102853). The per-handler try/catch stays as belt-and-suspenders.
      if (isSet('WordApi', '1.6')) {
        for (const evtHandlers of [
          doc.onParagraphChanged,
          doc.onParagraphAdded,
          doc.onParagraphDeleted,
        ]) {
          try {
            removers.push(evtHandlers.add(onDocChange));
          } catch {
            // This paragraph event isn't in the active requirement set — skip it.
          }
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
          handlers.onCommentAdded({
            ...(args.source !== undefined ? { source: args.source } : {}),
            ...(args.commentId !== undefined ? { commentId: args.commentId } : {}),
            ...(args.id !== undefined ? { id: args.id } : {}),
            ...(args.ids !== undefined ? { ids: args.ids } : {}),
          });
          return Promise.resolve();
        };
        try {
          removers.push(maybeComment.add(onComment));
        } catch {
          // Comment events unavailable — skip.
        }
      }

      await ctx.sync();
    }).catch(() => {
      // Word.run failed (no host / unsupported) — nothing registered, nothing to clean up.
    });

    return () => {
      if (unsubscribed) return; // idempotent: only the first call performs teardown.
      unsubscribed = true;
      removeSelection();
      // Single removal path: wait for registration to settle, then drain the removers once.
      void registration.then(() => removeWordHandlers(removers.splice(0)));
    };
  }
}

/** Remove Word object-model event handlers; per the typings, `.remove()` runs inside a sync batch. */
async function removeWordHandlers(handlers: Array<{ remove(): void }>): Promise<void> {
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
