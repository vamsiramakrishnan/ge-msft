import type { ActuationRequest, ContextRef } from '@ge/contracts';
import { isSet } from './capabilities-runtime.js';
import { chooseAnchorIndex, wordEdits } from './actuate-plan.js';
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

/**
 * Cap a bulk find/replace so one request can't rewrite an unbounded share of the document (the
 * registry's failure mode "too many hits" — the preview asks for clarification above the policy
 * limit; apply-time enforces the hard bound).
 */
export const MAX_FIND_REPLACEMENTS = 100;

/** Cap comment threads read into context so a heavily reviewed document can't blow the budget. */
export const MAX_READ_COMMENTS = 50;

/** One reply in a Word comment thread. */
export interface WordCommentReply {
  readonly authorName: string;
  readonly content: string;
}

/**
 * A comment thread read from the document: the comment, the text it is anchored on, and its
 * replies. All of it is untrusted document content — the runtime carries it as data.
 */
export interface WordComment {
  readonly id: string;
  readonly authorName: string;
  readonly content: string;
  readonly resolved: boolean;
  readonly anchorText: string;
  readonly replies: readonly WordCommentReply[];
}

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

/**
 * The outcome of an `apply-style` (ADR-0007). On success it carries the range's PRIOR style name so
 * the write is reversible via a `restore-style` inverse. `drift` degrades a stale content anchor to
 * a panel item, exactly like {@link TrackedChangeOutcome}.
 */
export type ApplyStyleOutcome =
  | { readonly status: 'applied'; readonly location: string; readonly priorStyle: string }
  | { readonly status: 'drift' };

/**
 * The outcome of an `insert-table` (ADR-0007): where the native table landed. An insertion has no
 * prior content, so reversibility is "delete the inserted table" — recorded as non-reversible
 * until a durable table handle is captured. `drift` is the usual stale-anchor degrade.
 */
export type InsertTableOutcome =
  | { readonly status: 'applied'; readonly location: string }
  | { readonly status: 'drift' }
  /** The paragraphs to replace run past {@link MAX_TABLE_REPLACE_PARAGRAPHS}: nothing changed. */
  | { readonly status: 'too_long'; readonly paragraphs: number }
  /** Paragraphs can only be replaced as a forced tracked change; this host cannot: nothing changed. */
  | { readonly status: 'untracked' }
  /** The run crosses something one delete would break: nothing changed. */
  | { readonly status: 'blocked'; readonly reason: string }
  /** The write was sent but failed part-way: check the document before trying again. */
  | { readonly status: 'unknown' };

/** A table replaces at most this many paragraphs, so a loose anchor cannot wipe a document. */
export const MAX_TABLE_REPLACE_PARAGRAPHS = 50;

/**
 * The outcome of an `insert-content-control` (ADR-0007). On success it carries the host-MINTED
 * control id (`cc.id`, read back post-sync) so the inverse deletes exactly THIS control — never a
 * re-resolved arbitrary one (the ADR-0007 §inverse-identity rule).
 */
export type InsertContentControlOutcome =
  | { readonly status: 'applied'; readonly location: string; readonly contentControlId: string }
  | { readonly status: 'drift' };

/**
 * The outcome of an `insert-hyperlink` (ADR-0007). On success it carries the range's PRIOR
 * hyperlink address so the write is reversible by restoring it. `drift` is the usual degrade.
 */
export type InsertHyperlinkOutcome =
  | { readonly status: 'applied'; readonly location: string; readonly priorHyperlink: string }
  | { readonly status: 'drift' };

/**
 * The outcome of a `find-replace` (ADR-0007): how many bounded hits were actually replaced. The
 * caller previews hit counts before writing; at apply-time the replacement is CAPPED at
 * {@link MAX_FIND_REPLACEMENTS} so a frequent term can't blow the blast radius.
 */
export type FindReplaceOutcome =
  | { readonly status: 'applied'; readonly replacedCount: number }
  | { readonly status: 'none' }; // search found nothing — nothing was written

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

  /** The text of the paragraph the selection (or the bare cursor) is in; undefined if unknown. */
  readCursorParagraph?(): Promise<string | undefined>;

  /** The whole body text (for the document context preview). */
  readBodyText(): Promise<string>;

  /** Non-empty body paragraphs with their built-in style (for heading-level derivation). */
  readParagraphs(): Promise<WordParagraph[]>;
  /** Host-side pagination; blanks retain their ordinal so cursors cannot skip content. */
  readParagraphPage?(offset: number, count: number): Promise<WordParagraph[]>;

  /**
   * Lazily read the body for `query` (ADR-0003 `search_document`): re-resolve the matches at
   * call-time via `body.search`, returning a bounded set of hits — each the matched text plus a
   * short surrounding-paragraph hint. Content-anchored (no stored offsets). Never throws; an
   * unsupported host / no hits yields `[]`.
   */
  searchText(query: string, matchCase: boolean): Promise<WordSearchHit[]>;

  /**
   * Read up to {@link MAX_READ_COMMENTS} comment threads — author, text, the anchored range's text
   * and the replies. Read-only. Never throws; a host without the comments API yields `[]`.
   */
  readComments(): Promise<WordComment[]>;

  /** Bring an addressable Word object into view without changing document content. */
  revealContext(ref: ContextRef): Promise<void>;

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

  /**
   * Tracked replace on the CURRENT selection — the fallback when a `suggest` anchor is the selected
   * text but `body.search` cannot find it (over its 255-character limit, or spanning paragraphs).
   * `matches` confirms, inside the same batch, that the live selection is still that text; otherwise
   * `drift` and nothing is written.
   */
  applyTrackedChangeToSelection(
    text: string,
    matches: (selectionText: string) => boolean,
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
    font?: TextFont,
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
   * ADR-0007 `apply-style`: set `styleName` on the resolved range — as a built-in style
   * (`styleBuiltIn`, WordApi 1.3) when `builtIn`, else the localized name (`style`, WordApi 1.1).
   * Anchored exactly like {@link insertText} (content anchor + `choose`, else the selection) and
   * CAPTURES the prior style first so the write is reversible.
   */
  applyStyle(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    styleName: string,
    builtIn: boolean,
    choose: ChooseHit,
  ): Promise<ApplyStyleOutcome>;

  /**
   * ADR-0007 `insert-table`: build a native table from a rectangular value grid at the selection or,
   * when `query` is given, after the chosen content-anchored hit (drift degrades). Anchored exactly
   * like {@link insertText}.
   */
  insertTable(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    rowCount: number,
    columnCount: number,
    values: readonly (readonly string[])[],
    choose: ChooseHit,
    /** Replace the paragraphs from the anchor's through the one containing this text. */
    replaceThrough?: string,
  ): Promise<InsertTableOutcome>;

  /**
   * ADR-0007 `insert-content-control`: wrap the selection or anchored range in a NEW content control
   * with an optional type/tag/title, and read back the host-MINTED control id for the inverse.
   */
  insertContentControl(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    controlType: string | undefined,
    tag: string | undefined,
    title: string | undefined,
    choose: ChooseHit,
  ): Promise<InsertContentControlOutcome>;

  /**
   * ADR-0007 `insert-hyperlink`: set the hyperlink on the selection or anchored range, CAPTURING the
   * prior address first so it is reversible. The URL is untrusted — the bridge screens it; this port
   * passes it to `range.hyperlink` as data.
   */
  insertHyperlink(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    url: string,
    choose: ChooseHit,
  ): Promise<InsertHyperlinkOutcome>;

  /**
   * ADR-0007 `find-replace`: replace every occurrence of `find` across the body with `replace`,
   * bounded at {@link MAX_FIND_REPLACEMENTS} hits. Returns how many hits were replaced; `none`
   * signals zero hits (nothing was written).
   */
  findReplace(
    find: string,
    replace: string,
    opts: { readonly matchCase: boolean; readonly matchWholeWord: boolean },
  ): Promise<FindReplaceOutcome>;

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
  /**
   * `trackEdits` (default on): every text edit is recorded as a Word tracked change, so a reviewer
   * accepts or rejects each one (product decision 2026-10-15). It is the one switch for all edits.
   */
  constructor(private readonly options: { readonly trackEdits?: boolean } = {}) {}

  /**
   * Runs one document edit as a tracked change when `trackEdits` is on, then puts the user's own
   * Track Changes setting back. WordApi 1.6: anchored edits must skip tracked deletions
   * (`withoutTrackedDeletions`), which needs 1.6, so tracking is never on without that filter.
   */
  private edit<T>(batch: (ctx: Word.RequestContext) => Promise<T>): Promise<T> {
    return Word.run(async (ctx) => {
      if (this.options.trackEdits === false || !isSet('WordApi', '1.6')) return batch(ctx);
      const doc = ctx.document;
      doc.load('changeTrackingMode');
      await ctx.sync();
      const prior = doc.changeTrackingMode;
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      try {
        return await batch(ctx);
      } finally {
        if (prior !== Word.ChangeTrackingMode.trackAll) {
          doc.changeTrackingMode = prior;
          // Never mask the edit's own outcome with a failed restore.
          await ctx.sync().catch(() => undefined);
        }
      }
    });
  }

  async readSelectionText(): Promise<string> {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      return (await reviewedText(ctx, sel)) ?? sel.text;
    });
  }

  async readCursorParagraph(): Promise<string | undefined> {
    try {
      return await Word.run(async (ctx) => {
        const paragraph = ctx.document.getSelection().paragraphs.getFirstOrNullObject();
        paragraph.load('text,isNullObject');
        await ctx.sync();
        return paragraph.isNullObject ? undefined : paragraph.text;
      });
    } catch {
      return undefined;
    }
  }

  async readBodyText(): Promise<string> {
    return Word.run(async (ctx) => {
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return (await reviewedText(ctx, body)) ?? body.text;
    });
  }

  async readParagraphs(): Promise<WordParagraph[]> {
    return Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items/text,items/styleBuiltIn');
      await ctx.sync();
      const texts = await reviewedTexts(ctx, paras.items);
      return paras.items
        .map((p, i) => ({ text: texts[i] ?? p.text, styleBuiltIn: String(p.styleBuiltIn) }))
        .filter((p) => p.text.trim().length > 0);
    });
  }

  async readParagraphPage(offset: number, count: number): Promise<WordParagraph[]> {
    return Word.run(async (ctx) => {
      const paragraphs = ctx.document.body.paragraphs;
      paragraphs.load({ select: 'text,styleBuiltIn', skip: offset, top: count });
      await ctx.sync();
      if (paragraphs.items.length > count) throw new Error('Word did not honor paged reads.');
      return paragraphs.items.map((p) => ({ text: p.text, styleBuiltIn: String(p.styleBuiltIn) }));
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
        await ctx.sync();
        // Bound the hits before loading a paragraph for each: a common word can match thousands.
        const live = (await withoutTrackedDeletions(ctx, results.items)).slice(0, MAX_SEARCH_HITS);
        const paras = live.map((r) => r.paragraphs.getFirstOrNullObject());
        for (const p of paras) p.load('text');
        await ctx.sync();
        // The paragraphs after a hit: a heading or a short line alone does not show the content the
        // task is about (live 2026-10-01: "Payment terms" matched, and the sentences under it never
        // reached the model on a 66-paragraph document).
        const following = await followingText(ctx, paras);

        return live.map((r, i) => {
          const para = paras[i];
          const own = para && !para.isNullObject ? para.text : undefined;
          // Only the added paragraphs are capped; the hit's own paragraph is kept whole.
          const paraText =
            own !== undefined && following[i]
              ? `${own}\n${following[i]!.slice(0, HIT_CONTEXT_CHARS)}`
              : own;
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

  async readComments(): Promise<WordComment[]> {
    // `Body.getComments`, `Comment.getRange` and `Comment.replies` → WordApi 1.4.
    if (!isSet('WordApi', '1.4')) return [];
    try {
      return await Word.run(async (ctx) => {
        const comments = ctx.document.body.getComments();
        comments.load('items/id,items/authorName,items/content,items/resolved');
        await ctx.sync();
        // Queue every thread's anchor text + replies, then read them back in ONE more sync, so the
        // cost is two round-trips whatever the comment count.
        const threads = comments.items.slice(0, MAX_READ_COMMENTS).map((comment) => {
          const range = comment.getRange();
          range.load('text');
          const replies = comment.replies;
          replies.load('items/authorName,items/content');
          return { comment, range, replies };
        });
        await ctx.sync();
        return threads.map(({ comment, range, replies }) => ({
          id: hostCommentId(comment.id),
          authorName: comment.authorName ?? '',
          content: comment.content ?? '',
          resolved: comment.resolved === true,
          anchorText: range.text ?? '',
          replies: replies.items.map((r) => ({
            authorName: r.authorName ?? '',
            content: r.content ?? '',
          })),
        }));
      });
    } catch {
      // Comments unavailable / host quirk — the read degrades to nothing, never throws.
      return [];
    }
  }

  async revealContext(ref: ContextRef): Promise<void> {
    if (ref.surface !== 'word') return;

    const paragraphIndex = wordParagraphIndex(ref);
    if (paragraphIndex !== undefined) {
      await this.revealParagraph(paragraphIndex);
      return;
    }

    const contentControlId = wordContentControlId(ref);
    if (contentControlId !== undefined) {
      await this.revealContentControl(contentControlId);
      return;
    }

    const commentId = wordCommentId(ref);
    if (commentId !== undefined) {
      await this.revealComment(commentId);
      return;
    }

    const matchText = ref.anchor?.matchText.trim();
    if (matchText) {
      await this.revealTextAnchor(matchText, ref.anchor?.contextHint);
      return;
    }

    if (ref.kind === 'selection' || ref.id === 'word:selection') {
      await Word.run(async (ctx) => {
        ctx.document.getSelection().select();
        await ctx.sync();
      });
    }
  }

  private async revealParagraph(index: number): Promise<void> {
    if (!Number.isInteger(index) || index < 1) return;
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items');
      await ctx.sync();
      const para = paras.items[index - 1];
      if (!para) return;
      para.getRange().select();
      await ctx.sync();
    });
  }

  private async revealTextAnchor(
    matchText: string,
    contextHint: string | undefined,
  ): Promise<void> {
    await Word.run(async (ctx) => {
      const results = ctx.document.body.search(matchText, { matchCase: false });
      results.load('items/text');
      await ctx.sync();
      const idx = chooseAnchorIndex(
        results.items.map((r) => r.text),
        contextHint,
      );
      const range = idx >= 0 ? results.items[idx] : undefined;
      if (!range) return;
      range.select();
      await ctx.sync();
    });
  }

  private async revealContentControl(contentControlId: string): Promise<void> {
    if (!isSet('WordApi', '1.3')) return;
    const id = Number(contentControlId);
    if (!Number.isInteger(id)) return;
    await Word.run(async (ctx) => {
      const cc = ctx.document.contentControls.getByIdOrNullObject(id);
      cc.load('isNullObject');
      await ctx.sync();
      if (cc.isNullObject) return;
      cc.select();
      await ctx.sync();
    });
  }

  private async revealComment(commentId: string): Promise<void> {
    if (!isSet('WordApi', '1.4')) return;
    await Word.run(async (ctx) => {
      const comments = ctx.document.body.getComments();
      comments.load('items/id');
      await ctx.sync();
      const comment = comments.items.find((c) => hostCommentId(c.id) === commentId);
      if (!comment) return;
      comment.getRange().select();
      await ctx.sync();
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
      const live = await withoutTrackedDeletions(ctx, results.items);

      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };

      await replaceTracked(ctx, range, range.text, text);
      await ctx.sync();
      return { status: 'applied', location: 'tracked-change' };
    });
  }

  async applyTrackedChangeToSelection(
    text: string,
    matches: (selectionText: string) => boolean,
  ): Promise<TrackedChangeOutcome> {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      if (!sel.text.trim() || !matches(sel.text)) return { status: 'drift' };
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await replaceTracked(ctx, sel, sel.text, text);
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
      const comment = comments.items.find((c) => hostCommentId(c.id) === commentId);
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
    font?: TextFont,
  ): Promise<InsertOutcome> {
    return this.edit(async (ctx) => {
      if (query === undefined) {
        // No anchor → insert at the current selection (replace its content with the new text).
        const sel = ctx.document.getSelection();
        applyFont(sel.insertText(text, Word.InsertLocation.replace), font);
        await ctx.sync();
        return { status: 'applied', location: 'selection', insertedText: text };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      // Read-then-write: re-resolve the anchor before inserting so a drifted finding degrades.
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      // Insert AFTER the anchored range (a direct edit appends to the matched content).
      applyFont(range.insertText(text, Word.InsertLocation.after), font);
      await ctx.sync();
      return { status: 'applied', location: 'insert-text', insertedText: text };
    });
  }

  async replaceSelection(text: string): Promise<ReplaceSelectionOutcome> {
    return this.edit(async (ctx) => {
      const sel = ctx.document.getSelection();
      // Capture the prior selection text BEFORE overwriting — this is the inverse's restore payload.
      sel.load('text');
      await ctx.sync();
      const priorText = sel.text;
      if (priorText.length === 0) return { status: 'empty' };
      // Word-level edits, as a suggest: a grammar fix by /replace-selection marked the whole
      // sentence deleted and re-inserted (UAT #17, regression run 2026-10-17).
      await replaceTracked(ctx, sel, priorText, text);
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
    return this.edit(async (ctx) => {
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        sel.insertOoxml(ooxml, Word.InsertLocation.replace);
        await ctx.sync();
        return { status: 'applied', location: 'selection' };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      // Block content (a paragraph or table) goes after the anchor's paragraph: inserted after the
      // matched text it merged into that sentence ("…may terminateSummary", live 2026-10-01).
      if (/<w:(?:p|tbl)[\s>/]/.test(ooxml)) {
        range.paragraphs.getLast().getRange('Whole').insertOoxml(ooxml, Word.InsertLocation.after);
      } else {
        range.insertOoxml(ooxml, Word.InsertLocation.after);
      }
      await ctx.sync();
      return { status: 'applied', location: 'insert-ooxml' };
    });
  }

  async fillContentControl(
    contentControlId: string,
    text: string,
  ): Promise<FillContentControlOutcome> {
    return this.edit(async (ctx) => {
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

  async applyStyle(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    styleName: string,
    builtIn: boolean,
    choose: ChooseHit,
  ): Promise<ApplyStyleOutcome> {
    return this.edit(async (ctx) => {
      // `styleBuiltIn` is the locale-portable name but needs WordApi 1.3; the localized `style`
      // is WordApi 1.1 — the bridge gates per variant, this batch writes whichever was planned.
      const writeStyle = (target: Word.Range): void => {
        if (builtIn) target.styleBuiltIn = styleName as Word.BuiltInStyleName;
        else target.style = styleName;
      };
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        sel.load('style');
        await ctx.sync();
        const priorStyle = sel.style;
        writeStyle(sel);
        await ctx.sync();
        return { status: 'applied', location: 'selection', priorStyle };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      // Read-then-write: re-resolve the anchor before styling so a drifted finding degrades.
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      range.load('style');
      await ctx.sync();
      const priorStyle = range.style;
      writeStyle(range);
      await ctx.sync();
      return { status: 'applied', location: 'apply-style', priorStyle };
    });
  }

  async insertTable(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    rowCount: number,
    columnCount: number,
    values: readonly (readonly string[])[],
    choose: ChooseHit,
    replaceThrough?: string,
  ): Promise<InsertTableOutcome> {
    return this.edit(async (ctx) => {
      if (query !== undefined && replaceThrough !== undefined) {
        // Deleting paragraphs is only allowed as a tracked change the user can reject: without
        // forced tracking (WordApi < 1.6, or tracking off) it would be permanent (review H2).
        if (this.options.trackEdits === false || !isSet('WordApi', '1.6'))
          return { status: 'untracked' };
        const starts = ctx.document.body.search(query, { matchCase: opts.matchCase });
        const ends = ctx.document.body.search(replaceThrough, { matchCase: opts.matchCase });
        starts.load('items/text');
        ends.load('items/text');
        await ctx.sync();
        const liveStarts = await withoutTrackedDeletions(ctx, starts.items);
        const start = liveStarts[choose(liveStarts.map((r) => r.text))];
        if (!start) return { status: 'drift' };
        // Compare whole paragraphs, not hits: the run ends at the NEAREST paragraph containing the
        // end text at or after the start's paragraph, even when both texts share one (review M1).
        const startPara = start.paragraphs.getFirst().getRange('Whole');
        const liveEnds = await withoutTrackedDeletions(ctx, ends.items);
        const endParas = liveEnds.map((end) => end.paragraphs.getLast().getRange('Whole'));
        const relations = endParas.map((para) => startPara.compareLocationWith(para));
        await ctx.sync();
        const endPara = endParas.find((_, i) =>
          ['Equal', 'Before', 'AdjacentBefore'].includes(String(relations[i]?.value)),
        );
        if (!endPara) return { status: 'drift' };
        const block = startPara.expandTo(endPara);
        block.paragraphs.load('items');
        block.tables.load('items');
        block.sections.load('items');
        block.contentControls.load('items/cannotDelete');
        await ctx.sync();
        const count = block.paragraphs.items.length;
        if (count > MAX_TABLE_REPLACE_PARAGRAPHS) return { status: 'too_long', paragraphs: count };
        // Nothing that one delete could break or half-remove (review M4).
        if (block.tables.items.length > 0) return { status: 'blocked', reason: 'a table' };
        if (block.sections.items.length > 1)
          return { status: 'blocked', reason: 'a section break' };
        if (block.contentControls.items.some((cc) => cc.cannotDelete))
          return { status: 'blocked', reason: 'a locked content control' };
        // One tracked edit: the table goes in before the run, then the run is deleted.
        block.insertTable(rowCount, columnCount, Word.InsertLocation.before, values as string[][]);
        block.delete();
        try {
          await ctx.sync();
        } catch {
          // Office does not roll back a queued insert when the delete fails: the outcome is unknown.
          return { status: 'unknown' };
        }
        return { status: 'applied', location: 'insert-table' };
      }
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        sel.insertTable(rowCount, columnCount, Word.InsertLocation.after, values as string[][]);
        await ctx.sync();
        return { status: 'applied', location: 'selection' };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      range.insertTable(rowCount, columnCount, Word.InsertLocation.after, values as string[][]);
      await ctx.sync();
      return { status: 'applied', location: 'insert-table' };
    });
  }

  async insertContentControl(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    controlType: string | undefined,
    tag: string | undefined,
    title: string | undefined,
    choose: ChooseHit,
  ): Promise<InsertContentControlOutcome> {
    return this.edit(async (ctx) => {
      // The `contentControlType` parameter itself arrived in WordApi 1.5 — omit it entirely on the
      // default rich-text path so a 1.1 host still gets a control. tag/title setters are WordApi 1.1.
      const controlTypeArg =
        controlType === undefined ? undefined : (controlType as Word.ContentControlType.richText);
      const wrap = (target: Word.Range): Word.ContentControl => {
        const cc = target.insertContentControl(controlTypeArg);
        if (tag !== undefined) cc.tag = tag;
        if (title !== undefined) cc.title = title;
        return cc;
      };
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        const cc = wrap(sel);
        // Read back the MINTED id post-sync: the inverse deletes exactly THIS control.
        cc.load('id');
        await ctx.sync();
        return {
          status: 'applied',
          location: `content-control:${cc.id}`,
          contentControlId: String(cc.id),
        };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      const cc = wrap(range);
      cc.load('id');
      await ctx.sync();
      return {
        status: 'applied',
        location: `content-control:${cc.id}`,
        contentControlId: String(cc.id),
      };
    });
  }

  async insertHyperlink(
    query: string | undefined,
    opts: { readonly matchCase: boolean },
    url: string,
    choose: ChooseHit,
  ): Promise<InsertHyperlinkOutcome> {
    return this.edit(async (ctx) => {
      if (query === undefined) {
        const sel = ctx.document.getSelection();
        sel.load('hyperlink');
        await ctx.sync();
        const priorHyperlink = sel.hyperlink;
        sel.hyperlink = url;
        await ctx.sync();
        return { status: 'applied', location: 'selection', priorHyperlink };
      }
      const results = ctx.document.body.search(query, { matchCase: opts.matchCase });
      results.load('items/text');
      await ctx.sync();
      // A tracked deletion still matches the search: never anchor an edit on struck-out text.
      const live = await withoutTrackedDeletions(ctx, results.items);
      const idx = choose(live.map((r) => r.text));
      const range = idx >= 0 ? live[idx] : undefined;
      if (!range) return { status: 'drift' };
      range.load('hyperlink');
      await ctx.sync();
      const priorHyperlink = range.hyperlink;
      range.hyperlink = url;
      await ctx.sync();
      return { status: 'applied', location: 'insert-hyperlink', priorHyperlink };
    });
  }

  async findReplace(
    find: string,
    replace: string,
    opts: { readonly matchCase: boolean; readonly matchWholeWord: boolean },
  ): Promise<FindReplaceOutcome> {
    return this.edit(async (ctx) => {
      const results = ctx.document.body.search(find, {
        matchCase: opts.matchCase,
        matchWholeWord: opts.matchWholeWord,
      });
      results.load('items/text');
      await ctx.sync();
      // Hard blast-radius bound: replace at most MAX_FIND_REPLACEMENTS hits per change.
      const targets = (await withoutTrackedDeletions(ctx, results.items)).slice(
        0,
        MAX_FIND_REPLACEMENTS,
      );
      // Last hit first: with Track Changes on, an earlier replacement in the same paragraph shifted
      // the later hits and the text landed mid-word ("th=>e Contoso-AI").
      for (const hit of [...targets].reverse()) {
        hit.insertText(replace, Word.InsertLocation.replace);
      }
      await ctx.sync();
      if (targets.length === 0) return { status: 'none' };
      return { status: 'applied', replacedCount: targets.length };
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

export function canRevealWordContext(ref: ContextRef): boolean {
  return (
    ref.surface === 'word' &&
    (ref.kind === 'selection' ||
      wordParagraphIndex(ref) !== undefined ||
      ref.id === 'word:selection' ||
      Boolean(ref.anchor?.matchText.trim()) ||
      wordContentControlId(ref) !== undefined ||
      wordCommentId(ref) !== undefined)
  );
}

function wordParagraphIndex(ref: ContextRef): number | undefined {
  const raw =
    prefixedValue(ref.id, 'word:paragraph:', 'paragraph:', 'para:') ??
    prefixedValue(ref.anchor?.locator, 'word:paragraph:', 'paragraph:', 'para:');
  if (!raw) return undefined;
  const index = Number(raw);
  return Number.isInteger(index) && index > 0 ? index : undefined;
}

function wordContentControlId(ref: ContextRef): string | undefined {
  return (
    prefixedValue(ref.id, 'word:cc:', 'word:content-control:') ??
    prefixedValue(
      ref.anchor?.locator,
      'cc:',
      'content-control:',
      'word:cc:',
      'word:content-control:',
    )
  );
}

/**
 * Each paragraph's text as it reads now, without tracked deletions. `Paragraph.text` keeps the
 * struck-out original next to its replacement ("GEGemini Enterprise"), so after a tracked edit the
 * model read its own change as broken and kept "fixing" it. `getReviewedText` is WordApi 1.4; an
 * older host, or a failed read, keeps the plain text.
 */
async function reviewedTexts(
  ctx: Word.RequestContext,
  paragraphs: Word.Paragraph[],
): Promise<Array<string | undefined>> {
  if (paragraphs.length === 0 || !isSet('WordApi', '1.4')) return [];
  if (!paragraphs.every((p) => typeof p.getReviewedText === 'function')) return [];
  try {
    const reviewed = paragraphs.map((p) => p.getReviewedText('Current'));
    await ctx.sync();
    return reviewed.map((r) => r.value);
  } catch {
    return [];
  }
}

/** Body or selection text without tracked deletions (see {@link reviewedTexts}); undefined = use `.text`. */
async function reviewedText(
  ctx: Word.RequestContext,
  target: Word.Body | Word.Range,
): Promise<string | undefined> {
  if (!isSet('WordApi', '1.4') || typeof target.getReviewedText !== 'function') return undefined;
  try {
    const reviewed = target.getReviewedText('Current');
    await ctx.sync();
    return reviewed.value;
  } catch {
    return undefined;
  }
}

/**
 * The hits that are live text. With Track Changes on, replaced text stays in the body as a tracked
 * deletion and `body.search` still finds it, so a re-run replaced it again ("GEGE") and a search
 * reported it as never changed. A hit touching a deletion is dropped. `Range.getTrackedChanges` is
 * WordApi 1.6; older hosts keep every hit, as before.
 */
async function withoutTrackedDeletions(
  ctx: Word.RequestContext,
  hits: Word.Range[],
): Promise<Word.Range[]> {
  if (hits.length === 0 || !isSet('WordApi', '1.6')) return hits;
  const changes = hits.map((hit) => {
    const tracked = hit.getTrackedChanges();
    tracked.load('items/type');
    return tracked;
  });
  await ctx.sync();
  return hits.filter((_, i) => !changes[i]!.items.some((change) => change.type === 'Deleted'));
}

/**
 * A host comment id as the string the rest of the add-in compares. The typings say `string`, but
 * Word on the web returns a number, so a strict `===` against the listed id never matched and
 * every reply failed as "comment gone" even with the id in hand.
 */
function hostCommentId(id: unknown): string {
  return String(id).trim();
}

function wordCommentId(ref: ContextRef): string | undefined {
  if (ref.kind !== 'comment') return undefined;
  return (
    prefixedValue(ref.id, 'word:comment:', 'comment:') ??
    prefixedValue(ref.anchor?.locator, 'comment:', 'word:comment:')
  );
}

function prefixedValue(value: string | undefined, ...prefixes: string[]): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      const rest = trimmed.slice(prefix.length).trim();
      return rest || undefined;
    }
  }
  return undefined;
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

/** Paragraphs read after each search hit, and the context they may add. */
const FOLLOWING_PARAGRAPHS = 2;
const HIT_CONTEXT_CHARS = 600;

/** The text of up to {@link FOLLOWING_PARAGRAPHS} paragraphs after each given paragraph. */
async function followingText(
  ctx: Word.RequestContext,
  paras: readonly Word.Paragraph[],
): Promise<string[]> {
  const texts = paras.map(() => [] as string[]);
  let current: Array<Word.Paragraph | undefined> = paras.map((p) =>
    p.isNullObject ? undefined : p,
  );
  for (let step = 0; step < FOLLOWING_PARAGRAPHS; step++) {
    const next = current.map((p) =>
      p && typeof p.getNextOrNullObject === 'function' ? p.getNextOrNullObject() : undefined,
    );
    if (next.every((n) => n === undefined)) break;
    for (const n of next) n?.load('text');
    await ctx.sync();
    current = next.map((n) => (n && !n.isNullObject ? n : undefined));
    current.forEach((n, i) => {
      if (n && n.text.trim()) texts[i]!.push(n.text.trim());
    });
  }
  return texts.map((t) => t.join('\n'));
}

/** The font of an `insert-text` (`bold=true font=Arial size=16`). */
export type TextFont = NonNullable<ActuationRequest['params']['textFont']>;

/** Apply the requested font to the range `insertText` returned; unset fields keep the document's. */
function applyFont(range: Word.Range, font: TextFont | undefined): void {
  if (!font) return;
  if (font.bold !== undefined) range.font.bold = font.bold;
  if (font.italic !== undefined) range.font.italic = font.italic;
  if (font.underline !== undefined) range.font.underline = font.underline ? 'Single' : 'None';
  if (font.name !== undefined) range.font.name = font.name;
  if (font.size !== undefined) range.font.size = font.size;
  if (font.color !== undefined) range.font.color = font.color;
}

/**
 * Write `text` over `range` (tracked changes on) as word-level edits when that marks fewer words,
 * so a grammar fix shows only the corrected words; otherwise replace the whole range as before.
 * The word ranges must line up with the text that was read, or the whole range is replaced.
 */
/**
 * Word-level edits of a multi-paragraph range against a new text with a different paragraph count.
 * Each edit is applied to the words it covers; an edit that runs into the next paragraph keeps the
 * paragraph mark: its text goes into the first paragraph and the later words are removed. False
 * (nothing written) when the words do not line up or a word-level edit would not help.
 */
async function editAcrossParagraphs(
  ctx: Word.RequestContext,
  pieces: Word.Range[],
  oldParagraphs: string[],
  text: string,
): Promise<boolean> {
  const plainWords = (s: string) => s.split(/\s+/).filter(Boolean);
  if (oldParagraphs.some((p) => /\t|\u00a0| {2}/.test(p.trim()))) return false;
  const edits = wordEdits(oldParagraphs.join(' '), text.replace(/\s+/g, ' '));
  if (!edits?.length) return false;
  const words = pieces.map((piece) => {
    const ranges = piece.getTextRanges([' '], true);
    ranges.load('items/text');
    return ranges;
  });
  await ctx.sync();
  // Global word index → (paragraph, word range).
  const at: Array<{ paragraph: number; range: Word.Range }> = [];
  for (const [i, ranges] of words.entries()) {
    const expected = plainWords(oldParagraphs[i]!);
    if (
      ranges.items.length !== expected.length ||
      !ranges.items.every((word, j) => word.text.trim() === expected[j])
    )
      return false;
    for (const range of ranges.items) at.push({ paragraph: i, range });
  }
  // Back to front, so an earlier edit never shifts a later word range.
  for (const edit of [...edits].reverse()) {
    const span = at.slice(edit.start, edit.end);
    const first = span[0]!;
    const head = span.filter((w) => w.paragraph === first.paragraph);
    const target =
      head.length === 1 ? first.range : first.range.expandTo(head[head.length - 1]!.range);
    for (const rest of span.slice(head.length).reverse())
      rest.range.insertText('', Word.InsertLocation.replace);
    target.insertText(edit.text, Word.InsertLocation.replace);
  }
  return true;
}

/** A text's paragraphs, without a trailing empty one. */
function paragraphsOf(text: string): string[] {
  const parts = text.split(/\r\n?|\n/);
  while (parts.length > 1 && parts.at(-1)!.trim() === '') parts.pop();
  return parts;
}

async function replaceTracked(
  ctx: Word.RequestContext,
  range: Word.Range,
  oldText: string,
  text: string,
): Promise<void> {
  // Several paragraphs: edit each one against the model's matching paragraph, so only the changed
  // words are tracked and the paragraph marks stay (real Word 2026-10-17: a two-paragraph
  // selection was deleted whole and re-inserted as one paragraph).
  const oldParagraphs = paragraphsOf(oldText);
  if (oldParagraphs.length > 1 && isSet('WordApi', '1.3')) {
    const newParagraphs = paragraphsOf(text);
    try {
      const pieces = range.split(['\r'], true, false);
      pieces.load('items/text');
      await ctx.sync();
      const aligned =
        pieces.items.length === oldParagraphs.length &&
        pieces.items.every((piece, i) => piece.text.trim() === oldParagraphs[i]!.trim());
      if (aligned && newParagraphs.length === oldParagraphs.length) {
        // Back to front, so an edit never shifts a later paragraph's range.
        for (let i = pieces.items.length - 1; i >= 0; i--) {
          const piece = pieces.items[i]!;
          if (piece.text.trim() !== newParagraphs[i]!.trim())
            await replaceTracked(ctx, piece, piece.text, newParagraphs[i]!.trim());
        }
        return;
      }
      // The model joined or split paragraphs (it sees a selection as one line): diff the words of
      // the whole selection and apply each change inside its own paragraph. Paragraph marks are
      // never touched (real Word 2026-10-17, second run).
      if (aligned && (await editAcrossParagraphs(ctx, pieces.items, oldParagraphs, text))) return;
    } catch {
      // The split and word reads write nothing. Track changes again (see below) and go on.
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }
  }
  // Only plain single-spaced text: a run across a tab, line break or paragraph mark would be
  // rejoined with one space.
  const plain = !/[\t\n\r\u00a0]| {2}/.test(oldText) && !oldText.includes(String.fromCharCode(11));
  const edits = plain && isSet('WordApi', '1.3') ? wordEdits(oldText, text) : undefined;
  if (edits?.length) {
    let words: Word.RangeCollection;
    try {
      words = range.getTextRanges([' '], true);
      words.load('items/text');
      await ctx.sync();
    } catch {
      // The split is a read: nothing was written, so the whole-range replace is still safe. Track
      // changes again: the failed batch may have dropped the mode set before it.
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      range.insertText(text, Word.InsertLocation.replace);
      return;
    }
    const expected = oldText.split(/\s+/).filter(Boolean);
    const aligned =
      words.items.length === expected.length &&
      words.items.every((word, i) => word.text.trim() === expected[i]);
    if (aligned) {
      // Back to front, so an earlier edit never shifts a later word range.
      for (const edit of [...edits].reverse()) {
        const first = words.items[edit.start]!;
        const target =
          edit.end - edit.start === 1 ? first : first.expandTo(words.items[edit.end - 1]!);
        target.insertText(edit.text, Word.InsertLocation.replace);
      }
      return;
    }
  }
  range.insertText(text, Word.InsertLocation.replace);
}
