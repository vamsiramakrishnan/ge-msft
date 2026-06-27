import { isSet } from './capabilities-runtime.js';
import type { WordSearchHit } from './capture.js';

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

/** Cap lazy `search_document` reads so a frequent term can't blow the per-turn budget. */
const MAX_SEARCH_HITS = 8;

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

/**
 * The outcome of a direct (non-tracked) content insert — ADR-0007 `insert-text` / `insert-ooxml`.
 * On success it carries `location` (where the write landed) plus the prior-state needed to reverse
 * the DIRECT edit: an insertion has no prior text, so reversibility is "delete the range that was
 * inserted", anchored back by the very content that was written (`insertedText` for plain text, or a
 * caller-supplied anchor for OOXML where the rendered text isn't known). `drift` degrades a stale
 * content anchor to a panel item, exactly like {@link TrackedChangeOutcome}.
 */
export type InsertOutcome =
  | {
      readonly status: 'applied';
      readonly location: string;
      /** Prior-state for the inverse: the text that was inserted (for re-finding the range to delete). */
      readonly insertedText?: string;
    }
  | { readonly status: 'drift' }; // anchor intended but no live hit

/**
 * The outcome of a `replace-selection` (ADR-0007). On success it carries the prior selection text so
 * the edit is reversible by writing `priorText` back over the new selection. `empty` signals that
 * nothing was selected (no range to replace) — the bridge fails closed rather than inserting at an
 * undefined location.
 */
export type ReplaceSelectionOutcome =
  | { readonly status: 'applied'; readonly location: string; readonly priorText: string }
  | { readonly status: 'empty' };

/**
 * The outcome of a `fill-content-control` (ADR-0007). On success it carries the content control's
 * PRIOR text so the fill is reversible by restoring it. `gone` signals the control with that id no
 * longer exists (a stale id is the content-control analogue of anchor drift), so the bridge degrades
 * to a panel item rather than throwing.
 */
export type FillContentControlOutcome =
  | { readonly status: 'applied'; readonly location: string; readonly priorText: string }
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
   * Lazily read the body for `query` (ADR-0003 `search_document`): re-resolve the matches at
   * call-time via `body.search`, returning a bounded set of hits — each the matched text plus a
   * short surrounding-paragraph hint. Content-anchored (no stored offsets). Never throws; an
   * unsupported host / no hits yields `[]`.
   */
  searchText(query: string, matchCase: boolean): Promise<WordSearchHit[]>;

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
   * ADR-0007 `insert-text`: insert plain `text` directly (NOT a tracked change). When `query` is
   * given, re-resolve it via `body.search` and let `choose` pick the read-back hit, inserting after
   * the chosen range (degrading to `drift` when there's no live hit) — the same content-anchoring
   * discipline as {@link applyTrackedChange}. When `query` is `undefined`, insert at the current
   * selection. Returns the prior-state (`insertedText`) the bridge records for the inverse.
   */
  insertText(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome>;

  /**
   * ADR-0007 `replace-selection`: replace the current selection's text with `text`, CAPTURING the
   * prior selection text first so the edit is reversible. `empty` when nothing is selected.
   */
  replaceSelection(text: string): Promise<ReplaceSelectionOutcome>;

  /**
   * ADR-0007 `insert-ooxml`: insert rich `ooxml` directly via `range.insertOoxml`, anchored exactly
   * like {@link insertText} (content anchor + `choose`, else the selection). The OOXML is untrusted
   * data passed straight to the host. Returns the anchor used so the bridge can record the inverse.
   */
  insertOoxml(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    ooxml: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome>;

  /**
   * ADR-0007 `fill-content-control`: populate the content control with `contentControlId`, CAPTURING
   * its prior text first so the fill is reversible. `gone` when no control with that id exists.
   */
  fillContentControl(contentControlId: string, text: string): Promise<FillContentControlOutcome>;

  /**
   * Attach a Word comment carrying `text` to the range matching `query` (re-resolved at
   * apply-time, like {@link applyTrackedChange}). Used for ADR-0003 comments-as-citations after a
   * tracked change applies. Best-effort: returns `{ ok: false }` (never throws) when the comments
   * API is unsupported or the anchor text is gone — the caller must NOT treat that as failing the
   * underlying change.
   */
  addComment(query: string, matchCase: boolean, text: string): Promise<{ ok: boolean }>;

  /**
   * Persist a durable-provenance custom XML part carrying `xml` into the document (BUILD-PLAN 1.6).
   * Word's durable metadata is OOXML custom XML, so the record lands as a `customXmlParts.add(xml)`
   * part keyed by the write's `changeId`. Best-effort: returns `{ ok: false }` (never throws) when
   * the API is unsupported, so a persistence failure can't fail the reversible write it accompanies.
   */
  persistProvenance(xml: string): Promise<{ ok: boolean }>;

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

  async searchText(query: string, matchCase: boolean): Promise<WordSearchHit[]> {
    const q = query.trim();
    if (!q) return [];
    try {
      return await Word.run(async (ctx) => {
        const results = ctx.document.body.search(q, { matchCase });
        // Load the match text plus its surrounding paragraph (the short contextHint). `body.search`
        // → WordApi 1.1; reading a result's `paragraphs` is broadly available, but we guard the
        // whole batch in try/catch so an older/quirky host degrades to `[]` rather than throwing.
        results.load('items/text');
        const paras = results.items.map((r) => r.paragraphs.getFirstOrNullObject());
        for (const p of paras) p.load('text');
        await ctx.sync();

        return results.items.slice(0, MAX_SEARCH_HITS).map((r, i) => {
          const para = paras[i];
          const paraText = para && !para.isNullObject ? para.text : undefined;
          const hint =
            paraText !== undefined && paraText.trim() && paraText.trim() !== r.text.trim()
              ? paraText
              : undefined;
          return hint !== undefined ? { text: r.text, contextHint: hint } : { text: r.text };
        });
      });
    } catch {
      // Search unavailable / host quirk — lazy read degrades to nothing, never throws.
      return [];
    }
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

  async addComment(query: string, matchCase: boolean, text: string): Promise<{ ok: boolean }> {
    // Feature-detect the comments API by requirement set, NOT property truthiness:
    // `Range.insertComment` → WordApi 1.4 (typings l.108627). On an older host we skip silently
    // so a missing citation comment never disturbs the already-applied tracked change.
    if (!isSet('WordApi', '1.4')) return { ok: false };
    try {
      return await Word.run(async (ctx) => {
        const results = ctx.document.body.search(query, { matchCase });
        results.load('items');
        // Re-resolve the anchor at apply-time: the comment lands on the live range, or nowhere.
        await ctx.sync();
        const range = results.items[0];
        if (!range) return { ok: false };
        range.insertComment(text);
        await ctx.sync();
        return { ok: true };
      });
    } catch {
      // Host quirk / comments unavailable — best-effort, log-and-continue.
      return { ok: false };
    }
  }

  async persistProvenance(xml: string): Promise<{ ok: boolean }> {
    // `CustomXmlPartCollection.add(xml)` → WordApi 1.4 (typings l.100826); the document's
    // `customXmlParts` getter is `WordApiHiddenDocument 1.4` (l.102901). Gate on WordApi 1.4 and
    // degrade silently on an older host so a missing provenance part never disturbs the write.
    if (!isSet('WordApi', '1.4')) return { ok: false };
    try {
      return await Word.run(async (ctx) => {
        ctx.document.customXmlParts.add(xml);
        await ctx.sync();
        return { ok: true };
      });
    } catch {
      // Custom XML parts unavailable / host quirk — best-effort, log-and-continue.
      return { ok: false };
    }
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

  async insertText(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome> {
    return Word.run(async (ctx) => {
      if (query === undefined) {
        // No anchor → insert at the current selection (replace its content with the new text).
        const sel = ctx.document.getSelection();
        sel.insertText(text, Word.InsertLocation.replace);
        await ctx.sync();
        return { status: 'applied', location: 'selection', insertedText: text };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      // Read-then-write: re-resolve the anchor before inserting so a drifted finding degrades.
      await ctx.sync();
      const idx = choose(results.items.map((r) => r.text));
      const range = idx >= 0 ? results.items[idx] : undefined;
      if (!range) return { status: 'drift' };
      // Insert AFTER the anchored range (a direct edit appends to the matched content).
      range.insertText(text, Word.InsertLocation.after);
      await ctx.sync();
      return { status: 'applied', location: 'insert-text', insertedText: text };
    });
  }

  async replaceSelection(text: string): Promise<ReplaceSelectionOutcome> {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      // Capture the prior selection text BEFORE overwriting — this is the inverse's restore payload.
      sel.load('text');
      await ctx.sync();
      const priorText = sel.text;
      if (priorText.length === 0) return { status: 'empty' };
      sel.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      return { status: 'applied', location: 'selection', priorText };
    });
  }

  async insertOoxml(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    ooxml: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome> {
    return Word.run(async (ctx) => {
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        sel.insertOoxml(ooxml, Word.InsertLocation.replace);
        await ctx.sync();
        return { status: 'applied', location: 'selection' };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      await ctx.sync();
      const idx = choose(results.items.map((r) => r.text));
      const range = idx >= 0 ? results.items[idx] : undefined;
      if (!range) return { status: 'drift' };
      range.insertOoxml(ooxml, Word.InsertLocation.after);
      await ctx.sync();
      return { status: 'applied', location: 'insert-ooxml' };
    });
  }

  async fillContentControl(
    contentControlId: string,
    text: string,
  ): Promise<FillContentControlOutcome> {
    return Word.run(async (ctx) => {
      // `getByIdOrNullObject` resolves the named container; a stale id yields a null object (no
      // throw), which we degrade to `gone` — the content-control analogue of anchor drift.
      const cc = ctx.document.contentControls.getByIdOrNullObject(Number(contentControlId));
      cc.load('text,isNullObject');
      await ctx.sync();
      if (cc.isNullObject) return { status: 'gone' };
      // Capture the prior text BEFORE replacing — the inverse restores it.
      const priorText = cc.text;
      cc.insertText(text, Word.InsertLocation.replace);
      await ctx.sync();
      return { status: 'applied', location: `content-control:${contentControlId}`, priorText };
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
