import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import { WORD_CAPABILITIES } from './capabilities.js';
import {
  headingLevel,
  wordDocumentToContext,
  wordSelectionToContext,
  type WordElement,
} from './capture.js';
import { chooseAnchorIndex, planTrackedChange } from './actuate-plan.js';

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
}
