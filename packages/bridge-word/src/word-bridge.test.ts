import { describe, it, expect, vi } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { WordBridge } from './word-bridge.js';
import type {
  ChooseHit,
  CommentReplyOutcome,
  TrackedChangeOutcome,
  WordHandlers,
  WordHost,
  WordParagraph,
} from './host-port.js';

/**
 * A fully in-memory {@link WordHost} — no Office.js. It models the host's search→insert batch so
 * the bridge's orchestration (anchor choice, drift degradation, outcome→result mapping, event
 * wiring) runs against it. The fake re-implements ONLY the host's observable contract: it offers
 * the read-back hit texts to the bridge's `choose` callback and writes only on the chosen hit.
 */
class FakeWordHost implements WordHost {
  selectionText = '';
  bodyText = '';
  paragraphs: WordParagraph[] = [];
  /** Body "search index": query → the hit texts the host would read back. */
  searchHits = new Map<string, string[]>();
  /** Comment ids that currently exist. */
  comments = new Set<string>();

  // Recorded effects, for assertions.
  readonly inserts: Array<{ query: string; matchCase: boolean; text: string; chosen: string }> = [];
  readonly replies: Array<{ commentId: string; reply: string; resolve: boolean }> = [];
  lastHandlers?: WordHandlers;
  unsubscribed = false;

  readSelectionText(): Promise<string> {
    return Promise.resolve(this.selectionText);
  }
  readBodyText(): Promise<string> {
    return Promise.resolve(this.bodyText);
  }
  readParagraphs(): Promise<WordParagraph[]> {
    return Promise.resolve(this.paragraphs);
  }

  applyTrackedChange(
    query: string,
    opts: { matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<TrackedChangeOutcome> {
    const hits = this.searchHits.get(query) ?? [];
    const idx = choose(hits);
    const chosen = idx >= 0 ? hits[idx] : undefined;
    if (chosen === undefined) return Promise.resolve({ status: 'drift' });
    this.inserts.push({ query, matchCase: opts.matchCase, text, chosen });
    return Promise.resolve({ status: 'applied', location: 'tracked-change' });
  }

  replyToComment(commentId: string, reply: string, resolve: boolean): Promise<CommentReplyOutcome> {
    if (!this.comments.has(commentId)) return Promise.resolve({ status: 'gone' });
    this.replies.push({ commentId, reply, resolve });
    return Promise.resolve({ status: 'replied', location: `comment:${commentId}` });
  }

  registerHandlers(handlers: WordHandlers): () => void {
    this.lastHandlers = handlers;
    return () => {
      this.unsubscribed = true;
    };
  }
}

function trackedChange(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'tracked-change', surface: 'word', params };
}

describe('WordBridge orchestration (against a fake host)', () => {
  describe('listContext', () => {
    it('lists the selection (live) and the whole document, previewing both', async () => {
      const host = new FakeWordHost();
      host.selectionText = 'available 99.5% of the time';
      host.bodyText = 'Section 5. Service Levels. The services are available 99.5% of the time.';
      const refs = await new WordBridge(host).listContext();

      expect(refs).toHaveLength(2);
      expect(refs[0]).toMatchObject({ id: 'word:selection', kind: 'selection', live: true });
      expect(refs[1]).toMatchObject({ id: 'word:document', kind: 'document' });
      expect(refs[1]?.live).toBeUndefined();
    });

    it('omits the selection chip when the selection is blank', async () => {
      const host = new FakeWordHost();
      host.selectionText = '   \n  ';
      host.bodyText = 'body text';
      const refs = await new WordBridge(host).listContext();
      expect(refs).toHaveLength(1);
      expect(refs[0]?.kind).toBe('document');
    });
  });

  describe('resolveContext', () => {
    it('resolves a selection ref to a single live text part', async () => {
      const host = new FakeWordHost();
      host.selectionText = 'hello world';
      const ctx = await new WordBridge(host).resolveContext({
        id: 'word:selection',
        kind: 'selection',
        surface: 'word',
        title: 'Selection',
      });
      expect(ctx).toHaveLength(1);
      expect(ctx[0]).toMatchObject({
        ref: { kind: 'selection', live: true },
        value: { as: 'text' },
      });
    });

    it('resolves the document by mapping paragraphs (with heading levels) to context', async () => {
      const host = new FakeWordHost();
      host.paragraphs = [
        { text: 'Availability', styleBuiltIn: 'Heading2' },
        { text: 'The services are available 99.5% of the time.', styleBuiltIn: 'Normal' },
        { text: '   ', styleBuiltIn: 'Normal' },
      ];
      const ctx = await new WordBridge(host).resolveContext({
        id: 'word:document',
        kind: 'document',
        surface: 'word',
        title: 'Whole document',
      });
      expect(ctx.length).toBeGreaterThan(0);
    });
  });

  describe('actuate tracked-change', () => {
    it('applies a matching anchor and returns ok with the location and same changeId', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['intro 99.5%', 'Availability: 99.5% uptime']);

      const req = trackedChange(
        { text: '99.9%', target: { matchText: '99.5%', contextHint: 'Availability' } },
        'chg-42',
      );
      const res = await new WordBridge(host).actuate(req);

      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-42'),
        kind: 'tracked-change',
        location: 'tracked-change',
      });
      // changeId propagated, not re-minted.
      expect(res.changeId).toBe(req.changeId);
      // Wrote on the contextHint-matching hit, case-insensitively, exactly once.
      expect(host.inserts).toEqual([
        { query: '99.5%', matchCase: false, text: '99.9%', chosen: 'Availability: 99.5% uptime' },
      ]);
    });

    it('picks the first hit when no contextHint is given', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('SLA', ['first SLA', 'second SLA']);
      const res = await new WordBridge(host).actuate(
        trackedChange({ text: 'service level', target: { matchText: 'SLA' } }),
      );
      expect(res.ok).toBe(true);
      expect(host.inserts[0]?.chosen).toBe('first SLA');
    });

    it('degrades to a panel item (ok:false, anchor_drift, degraded) when the text is gone', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', []); // 0 matches → drift

      const req = trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }, 'chg-7');
      const res = await new WordBridge(host).actuate(req);

      expect(res).toMatchObject({
        ok: false,
        changeId: asChangeId('chg-7'),
        kind: 'tracked-change',
        degraded: true,
        error: { code: 'anchor_drift' },
      });
      // Crucially, NOTHING was written to any range.
      expect(host.inserts).toHaveLength(0);
    });

    it('degrades when the contextHint matches nothing and there are no hits at all', async () => {
      const host = new FakeWordHost();
      // hits exist but the query key differs → simulate a search that found nothing
      const res = await new WordBridge(host).actuate(
        trackedChange({ text: 'x', target: { matchText: 'absent', contextHint: 'nope' } }),
      );
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('anchor_drift');
      expect(host.inserts).toHaveLength(0);
    });

    it('rejects a tracked-change with no matchText anchor before touching the host', async () => {
      const host = new FakeWordHost();
      const insertSpy = vi.spyOn(host, 'applyTrackedChange');
      const res = await new WordBridge(host).actuate(trackedChange({ text: 'x' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
      expect(insertSpy).not.toHaveBeenCalled();
    });
  });

  describe('actuate comment-reply', () => {
    it('replies and resolves an existing comment, returning ok with a comment location', async () => {
      const host = new FakeWordHost();
      host.comments.add('cmt-1');
      const res = await new WordBridge(host).actuate({
        changeId: asChangeId('chg-c'),
        kind: 'comment-reply',
        surface: 'word',
        params: { text: 'done', target: { commentId: 'cmt-1' }, resolveComment: true },
      });
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-c'),
        kind: 'comment-reply',
        location: 'comment:cmt-1',
      });
      expect(host.replies).toEqual([{ commentId: 'cmt-1', reply: 'done', resolve: true }]);
    });

    it('degrades when the comment is gone', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate({
        changeId: asChangeId('chg-c'),
        kind: 'comment-reply',
        surface: 'word',
        params: { text: 'done', target: { commentId: 'missing' } },
      });
      expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'comment_gone' } });
    });

    it('rejects a comment-reply with no commentId', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate({
        changeId: asChangeId('chg-c'),
        kind: 'comment-reply',
        surface: 'word',
        params: { text: 'done' },
      });
      expect(res).toMatchObject({ ok: false, error: { code: 'no_comment' } });
    });
  });

  describe('actuate unsupported', () => {
    it('returns an unsupported error for kinds Word cannot do', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate({
        changeId: asChangeId('chg-x'),
        kind: 'write-cells',
        surface: 'word',
        params: { cells: [['1']] },
      });
      expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    });
  });

  describe('watch', () => {
    it('wires the host handlers and emits mapped HostEvents with correct origin', () => {
      const host = new FakeWordHost();
      const events: HostEvent[] = [];
      const unsub = new WordBridge(host).watch((e) => events.push(e));
      const handlers = host.lastHandlers;
      expect(handlers).toBeDefined();

      handlers?.onSelectionChanged();
      handlers?.onDocumentChanged({ source: 'Local' });
      handlers?.onDocumentChanged({ source: 'Remote' });
      handlers?.onCommentAdded({ source: 'Remote', commentId: 'cmt-9' });

      expect(events).toEqual([
        { type: 'selection-changed', surface: 'word', origin: 'local' },
        { type: 'document-changed', surface: 'word', origin: 'local' },
        { type: 'document-changed', surface: 'word', origin: 'remote' },
        { type: 'comment-added', surface: 'word', origin: 'remote', commentId: 'cmt-9' },
      ]);

      unsub();
      expect(host.unsubscribed).toBe(true);
    });

    it('falls back across commentId / id / ids and drops comment events without any id', () => {
      const host = new FakeWordHost();
      const events: HostEvent[] = [];
      new WordBridge(host).watch((e) => events.push(e));
      const h = host.lastHandlers;

      h?.onCommentAdded({ id: 'via-id' });
      h?.onCommentAdded({ ids: ['via-ids'] });
      h?.onCommentAdded({}); // no id at all → dropped

      expect(events).toEqual([
        { type: 'comment-added', surface: 'word', origin: 'local', commentId: 'via-id' },
        { type: 'comment-added', surface: 'word', origin: 'local', commentId: 'via-ids' },
      ]);
    });
  });
});
