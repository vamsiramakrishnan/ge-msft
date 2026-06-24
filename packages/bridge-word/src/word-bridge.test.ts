import { describe, it, expect, vi } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { WordBridge } from './word-bridge.js';
import { DocStateSnapshotSchema } from '@ge/contracts';
import type { WordSearchHit } from './capture.js';
import type {
  ChooseHit,
  CommentReplyOutcome,
  FillContentControlOutcome,
  InsertOutcome,
  ReplaceSelectionOutcome,
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
  /** Lazy-read search index for `searchText` (ADR-0003): query → re-resolved hits. */
  textHits = new Map<string, WordSearchHit[]>();
  /** Comment ids that currently exist. */
  comments = new Set<string>();
  /** Content controls that currently exist: id → current text. */
  contentControls = new Map<string, string>();

  // Recorded effects, for assertions.
  readonly inserts: Array<{ query: string; matchCase: boolean; text: string; chosen: string }> = [];
  readonly replies: Array<{ commentId: string; reply: string; resolve: boolean }> = [];
  readonly addedComments: Array<{ query: string; matchCase: boolean; text: string }> = [];
  /** Direct text/ooxml inserts (ADR-0007), for assertions. */
  readonly directInserts: Array<{
    query?: string;
    text?: string;
    ooxml?: string;
    chosen?: string;
  }> = [];
  /** replace-selection writes (ADR-0007), for assertions. */
  readonly selectionReplaces: Array<{ text: string; priorText: string }> = [];
  /** fill-content-control writes (ADR-0007), for assertions. */
  readonly filledControls: Array<{ id: string; text: string; priorText: string }> = [];
  /** Durable provenance XML parts persisted via the port (BUILD-PLAN 1.6). */
  readonly persistedProvenance: string[] = [];
  /** When true, the next addComment reports failure (unsupported / anchor gone). */
  commentFails = false;
  /** When true, persistProvenance reports failure (API unsupported). */
  provenanceFails = false;
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

  searchText(query: string, _matchCase: boolean): Promise<WordSearchHit[]> {
    return Promise.resolve(this.textHits.get(query) ?? []);
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

  addComment(query: string, matchCase: boolean, text: string): Promise<{ ok: boolean }> {
    if (this.commentFails) return Promise.resolve({ ok: false });
    this.addedComments.push({ query, matchCase, text });
    return Promise.resolve({ ok: true });
  }

  persistProvenance(xml: string): Promise<{ ok: boolean }> {
    if (this.provenanceFails) return Promise.resolve({ ok: false });
    this.persistedProvenance.push(xml);
    return Promise.resolve({ ok: true });
  }

  replyToComment(commentId: string, reply: string, resolve: boolean): Promise<CommentReplyOutcome> {
    if (!this.comments.has(commentId)) return Promise.resolve({ status: 'gone' });
    this.replies.push({ commentId, reply, resolve });
    return Promise.resolve({ status: 'replied', location: `comment:${commentId}` });
  }

  insertText(
    query: string | undefined,
    opts: { matchCase: boolean },
    text: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome> {
    if (query === undefined) {
      this.directInserts.push({ text });
      return Promise.resolve({ status: 'applied', location: 'selection', insertedText: text });
    }
    const hits = this.searchHits.get(query) ?? [];
    const idx = choose(hits);
    const chosen = idx >= 0 ? hits[idx] : undefined;
    if (chosen === undefined) return Promise.resolve({ status: 'drift' });
    void opts;
    this.directInserts.push({ query, text, chosen });
    return Promise.resolve({ status: 'applied', location: 'insert-text', insertedText: text });
  }

  replaceSelection(text: string): Promise<ReplaceSelectionOutcome> {
    const priorText = this.selectionText;
    if (priorText.length === 0) return Promise.resolve({ status: 'empty' });
    this.selectionReplaces.push({ text, priorText });
    this.selectionText = text;
    return Promise.resolve({ status: 'applied', location: 'selection', priorText });
  }

  insertOoxml(
    query: string | undefined,
    opts: { matchCase: boolean },
    ooxml: string,
    choose: ChooseHit,
  ): Promise<InsertOutcome> {
    if (query === undefined) {
      this.directInserts.push({ ooxml });
      return Promise.resolve({ status: 'applied', location: 'selection' });
    }
    const hits = this.searchHits.get(query) ?? [];
    const idx = choose(hits);
    const chosen = idx >= 0 ? hits[idx] : undefined;
    if (chosen === undefined) return Promise.resolve({ status: 'drift' });
    void opts;
    this.directInserts.push({ query, ooxml, chosen });
    return Promise.resolve({ status: 'applied', location: 'insert-ooxml' });
  }

  fillContentControl(contentControlId: string, text: string): Promise<FillContentControlOutcome> {
    const priorText = this.contentControls.get(contentControlId);
    if (priorText === undefined) return Promise.resolve({ status: 'gone' });
    this.filledControls.push({ id: contentControlId, text, priorText });
    this.contentControls.set(contentControlId, text);
    return Promise.resolve({
      status: 'applied',
      location: `content-control:${contentControlId}`,
      priorText,
    });
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

function addComment(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-comment', surface: 'word', params };
}

function insertTextReq(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-text', surface: 'word', params };
}

function replaceSelectionReq(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'replace-selection', surface: 'word', params };
}

function insertOoxmlReq(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-ooxml', surface: 'word', params };
}

function fillCcReq(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'fill-content-control', surface: 'word', params };
}

const PROVENANCE: ActuationRequest['provenance'] = {
  agentId: 'review@v1',
  identity: 'v.k@acme',
  timestamp: '2026-06-22T00:00:00Z',
  contentHash: 'h',
  sources: [{ title: 'SLA Policy' }],
};

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
        // This request carried no provenance payload → the write is flagged unattributed, so an
        // unprovenanced change is never mistaken for an attributed one.
        provenanceMissing: true,
      });
      // changeId propagated, not re-minted.
      expect(res.changeId).toBe(req.changeId);
      // Wrote on the contextHint-matching hit, case-insensitively, exactly once.
      expect(host.inserts).toEqual([
        { query: '99.5%', matchCase: false, text: '99.9%', chosen: 'Availability: 99.5% uptime' },
      ]);
    });

    it('flags provenanceDropped when the change lands but durable provenance fails to persist', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      host.provenanceFails = true; // host can't write the custom XML part.

      const req: ActuationRequest = {
        ...trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }, 'chg-prov'),
        provenance: PROVENANCE,
      };
      const res = await new WordBridge(host).actuate(req);

      // The reversible write still succeeded — the drop is surfaced, not fatal.
      expect(res.ok).toBe(true);
      expect(res.provenanceDropped).toBe(true);
    });

    it('does not flag provenanceDropped when persistence succeeds', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      const req: ActuationRequest = {
        ...trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }, 'chg-ok'),
        provenance: PROVENANCE,
      };
      const res = await new WordBridge(host).actuate(req);
      expect(res.ok).toBe(true);
      expect(res.provenanceDropped).toBeUndefined();
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

    it('adds a citation comment (formatted from sources) after a successful change', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      const req = trackedChange({
        text: '99.9%',
        target: { matchText: '99.5%' },
        sources: [{ title: 'SLA Policy', uri: 'https://acme/sla' }, { title: 'Uptime Memo' }],
      });
      const res = await new WordBridge(host).actuate(req);

      expect(res.ok).toBe(true);
      expect(host.addedComments).toEqual([
        {
          query: '99.5%',
          matchCase: false,
          text: 'SLA Policy (https://acme/sla)\nUptime Memo',
        },
      ]);
    });

    it('prefers provenance.sources over params.sources for the citation', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('SLA', ['the SLA']);
      const res = await new WordBridge(host).actuate({
        changeId: asChangeId('chg-p'),
        kind: 'tracked-change',
        surface: 'word',
        params: {
          text: 'service level',
          target: { matchText: 'SLA' },
          sources: [{ title: 'Fallback' }],
        },
        provenance: {
          agentId: 'review@v1',
          identity: 'v.k@acme',
          timestamp: '2026-06-22T00:00:00Z',
          contentHash: 'h',
          sources: [{ title: 'Preferred', uri: 'https://acme/preferred' }],
        },
      });
      expect(res.ok).toBe(true);
      expect(host.addedComments[0]?.text).toBe('Preferred (https://acme/preferred)');
    });

    it('still reports the change applied when adding the comment fails (best-effort)', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      host.commentFails = true;
      const res = await new WordBridge(host).actuate(
        trackedChange({
          text: '99.9%',
          target: { matchText: '99.5%' },
          sources: [{ title: 'SLA Policy' }],
        }),
      );
      // The change is applied even though the comment could not be attached.
      expect(res.ok).toBe(true);
      expect(host.inserts).toHaveLength(1);
      expect(host.addedComments).toHaveLength(0);
    });

    it('adds no comment when there are no sources', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      const commentSpy = vi.spyOn(host, 'addComment');
      const res = await new WordBridge(host).actuate(
        trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }),
      );
      expect(res.ok).toBe(true);
      expect(commentSpy).not.toHaveBeenCalled();
      expect(host.addedComments).toHaveLength(0);
    });

    it('adds no comment when the change degrades (drift)', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', []); // drift
      const commentSpy = vi.spyOn(host, 'addComment');
      const res = await new WordBridge(host).actuate(
        trackedChange({
          text: '99.9%',
          target: { matchText: '99.5%' },
          sources: [{ title: 'SLA Policy' }],
        }),
      );
      expect(res.ok).toBe(false);
      expect(res.error?.code).toBe('anchor_drift');
      expect(commentSpy).not.toHaveBeenCalled();
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
        provenanceMissing: true, // request carried no provenance payload (see tracked-change test)
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

  describe('actuate add-comment (ADR-0004)', () => {
    it('adds a content-anchored comment on a hit and returns ok with the same changeId', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate(
        addComment({ target: { matchText: '99.5%' }, text: 'Unsourced claim' }, 'chg-9'),
      );
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-9'),
        kind: 'add-comment',
        location: 'comment',
        provenanceMissing: true, // request carried no provenance payload (see tracked-change test)
      });
      expect(host.addedComments).toEqual([
        { query: '99.5%', matchCase: false, text: 'Unsourced claim' },
      ]);
    });

    it('degrades to a panel item (anchor_drift) when the anchor is gone', async () => {
      const host = new FakeWordHost();
      host.commentFails = true; // simulate the anchor text no longer present
      const res = await new WordBridge(host).actuate(
        addComment({ target: { matchText: 'gone' }, text: 'note' }, 'chg-d'),
      );
      expect(res).toMatchObject({
        ok: false,
        changeId: asChangeId('chg-d'),
        kind: 'add-comment',
        degraded: true,
        error: { code: 'anchor_drift' },
      });
    });

    it('rejects an add-comment with no matchText anchor before touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'addComment');
      const res = await new WordBridge(host).actuate(addComment({ text: 'note' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
      expect(spy).not.toHaveBeenCalled();
    });

    it('rejects an add-comment with empty text', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate(
        addComment({ target: { matchText: 'x' }, text: '   ' }),
      );
      expect(res).toMatchObject({ ok: false, error: { code: 'no_text' } });
    });
  });

  describe('durable provenance persistence (BUILD-PLAN 1.6)', () => {
    it('persists a provenance custom-XML part after a successful tracked change', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      const res = await new WordBridge(host).actuate({
        ...trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }, 'chg-prov'),
        provenance: PROVENANCE,
      });
      expect(res.ok).toBe(true);
      expect(host.persistedProvenance).toHaveLength(1);
      expect(host.persistedProvenance[0]).toContain('key="ge:prov:chg-prov"');
      expect(host.persistedProvenance[0]).toContain('agentId="review@v1"');
    });

    it('persists provenance after add-comment and comment-reply', async () => {
      const host = new FakeWordHost();
      host.comments.add('cmt-1');
      const bridge = new WordBridge(host);
      await bridge.actuate({
        ...addComment({ target: { matchText: 'x' }, text: 'note' }, 'chg-a'),
        provenance: PROVENANCE,
      });
      await bridge.actuate({
        changeId: asChangeId('chg-r'),
        kind: 'comment-reply',
        surface: 'word',
        params: { text: 'done', target: { commentId: 'cmt-1' } },
        provenance: PROVENANCE,
      });
      expect(host.persistedProvenance.map((x) => x.match(/key="([^"]+)"/)?.[1])).toEqual([
        'ge:prov:chg-a',
        'ge:prov:chg-r',
      ]);
    });

    it('skips persistence when the request carries no provenance', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      const res = await new WordBridge(host).actuate(
        trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }),
      );
      expect(res.ok).toBe(true);
      expect(host.persistedProvenance).toHaveLength(0);
    });

    it('does not persist provenance when the write degrades (drift)', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', []); // drift
      const res = await new WordBridge(host).actuate({
        ...trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }),
        provenance: PROVENANCE,
      });
      expect(res.ok).toBe(false);
      expect(host.persistedProvenance).toHaveLength(0);
    });

    it('still reports the write applied when provenance persistence fails (best-effort)', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['Availability: 99.5% uptime']);
      host.provenanceFails = true;
      const res = await new WordBridge(host).actuate({
        ...trackedChange({ text: '99.9%', target: { matchText: '99.5%' } }),
        provenance: PROVENANCE,
      });
      expect(res.ok).toBe(true);
      expect(host.persistedProvenance).toHaveLength(0);
    });
  });

  describe('actuate insert-text (ADR-0007)', () => {
    it('inserts at the current selection when no anchor is given, returning ok + same changeId', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate(insertTextReq({ text: 'hello' }, 'chg-i'));
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-i'),
        kind: 'insert-text',
        location: 'selection',
        provenanceMissing: true,
      });
      expect(host.directInserts).toEqual([{ text: 'hello' }]);
    });

    it('inserts at a content anchor, picking the contextHint hit', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', ['intro 99.5%', 'Availability: 99.5% uptime']);
      const res = await new WordBridge(host).actuate(
        insertTextReq({
          text: ' (revised)',
          target: { matchText: '99.5%', contextHint: 'Availability' },
        }),
      );
      expect(res.ok).toBe(true);
      expect(res.location).toBe('insert-text');
      expect(host.directInserts).toEqual([
        { query: '99.5%', text: ' (revised)', chosen: 'Availability: 99.5% uptime' },
      ]);
    });

    it('degrades (anchor_drift) when an intended anchor is gone, writing nothing', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('99.5%', []); // drift
      const res = await new WordBridge(host).actuate(
        insertTextReq({ text: 'x', target: { matchText: '99.5%' } }, 'chg-d'),
      );
      expect(res).toMatchObject({
        ok: false,
        changeId: asChangeId('chg-d'),
        kind: 'insert-text',
        degraded: true,
        error: { code: 'anchor_drift' },
      });
      expect(host.directInserts).toHaveLength(0);
    });

    it('fails closed on empty text before touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'insertText');
      const res = await new WordBridge(host).actuate(insertTextReq({ text: '' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_text' } });
      expect(spy).not.toHaveBeenCalled();
    });

    it('persists durable provenance after a successful insert', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate({
        ...insertTextReq({ text: 'hi' }, 'chg-ip'),
        provenance: PROVENANCE,
      });
      expect(res.ok).toBe(true);
      expect(host.persistedProvenance).toHaveLength(1);
    });
  });

  describe('actuate replace-selection (ADR-0007)', () => {
    it('captures prior text and replaces the selection, returning ok + same changeId', async () => {
      const host = new FakeWordHost();
      host.selectionText = 'old value';
      const res = await new WordBridge(host).actuate(
        replaceSelectionReq({ text: 'new value' }, 'chg-r'),
      );
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-r'),
        kind: 'replace-selection',
        location: 'selection',
        provenanceMissing: true,
      });
      // The prior text was captured (for the inverse) before the overwrite.
      expect(host.selectionReplaces).toEqual([{ text: 'new value', priorText: 'old value' }]);
    });

    it('degrades (no_selection) when nothing is selected', async () => {
      const host = new FakeWordHost();
      host.selectionText = '';
      const res = await new WordBridge(host).actuate(replaceSelectionReq({ text: 'x' }, 'chg-e'));
      expect(res).toMatchObject({
        ok: false,
        changeId: asChangeId('chg-e'),
        kind: 'replace-selection',
        degraded: true,
        error: { code: 'no_selection' },
      });
      expect(host.selectionReplaces).toHaveLength(0);
    });

    it('fails closed on empty text before touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'replaceSelection');
      const res = await new WordBridge(host).actuate(replaceSelectionReq({ text: '' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_text' } });
      expect(spy).not.toHaveBeenCalled();
    });
  });

  describe('actuate insert-ooxml (ADR-0007)', () => {
    it('inserts ooxml at the selection when no anchor is given', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate(insertOoxmlReq({ ooxml: '<w:p/>' }, 'chg-o'));
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-o'),
        kind: 'insert-ooxml',
        location: 'selection',
        provenanceMissing: true,
      });
      expect(host.directInserts).toEqual([{ ooxml: '<w:p/>' }]);
    });

    it('inserts ooxml at a content anchor', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('Summary', ['Summary section']);
      const res = await new WordBridge(host).actuate(
        insertOoxmlReq({ ooxml: '<w:tbl/>', target: { matchText: 'Summary' } }),
      );
      expect(res.ok).toBe(true);
      expect(res.location).toBe('insert-ooxml');
      expect(host.directInserts).toEqual([
        { query: 'Summary', ooxml: '<w:tbl/>', chosen: 'Summary section' },
      ]);
    });

    it('degrades (anchor_drift) when the anchor is gone', async () => {
      const host = new FakeWordHost();
      host.searchHits.set('Summary', []);
      const res = await new WordBridge(host).actuate(
        insertOoxmlReq({ ooxml: '<w:p/>', target: { matchText: 'Summary' } }, 'chg-od'),
      );
      expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'anchor_drift' } });
      expect(host.directInserts).toHaveLength(0);
    });

    it('fails closed on empty ooxml before touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'insertOoxml');
      const res = await new WordBridge(host).actuate(insertOoxmlReq({ ooxml: '' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_ooxml' } });
      expect(spy).not.toHaveBeenCalled();
    });
  });

  describe('actuate fill-content-control (ADR-0007)', () => {
    it('captures prior text and fills the control, returning ok + same changeId', async () => {
      const host = new FakeWordHost();
      host.contentControls.set('42', 'placeholder');
      const res = await new WordBridge(host).actuate(
        fillCcReq({ target: { contentControlId: '42' }, text: 'filled' }, 'chg-f'),
      );
      expect(res).toEqual({
        ok: true,
        changeId: asChangeId('chg-f'),
        kind: 'fill-content-control',
        location: 'content-control:42',
        provenanceMissing: true,
      });
      expect(host.filledControls).toEqual([{ id: '42', text: 'filled', priorText: 'placeholder' }]);
    });

    it('degrades (content_control_gone) when the control no longer exists', async () => {
      const host = new FakeWordHost();
      const res = await new WordBridge(host).actuate(
        fillCcReq({ target: { contentControlId: '99' }, text: 'x' }, 'chg-g'),
      );
      expect(res).toMatchObject({
        ok: false,
        changeId: asChangeId('chg-g'),
        kind: 'fill-content-control',
        degraded: true,
        error: { code: 'content_control_gone' },
      });
      expect(host.filledControls).toHaveLength(0);
    });

    it('fails closed on a missing contentControlId before touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'fillContentControl');
      const res = await new WordBridge(host).actuate(fillCcReq({ text: 'x' }));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_content_control' } });
      expect(spy).not.toHaveBeenCalled();
    });

    it('fails closed on empty text', async () => {
      const host = new FakeWordHost();
      host.contentControls.set('42', 'placeholder');
      const res = await new WordBridge(host).actuate(
        fillCcReq({ target: { contentControlId: '42' }, text: '' }),
      );
      expect(res).toMatchObject({ ok: false, error: { code: 'no_text' } });
      expect(host.filledControls).toHaveLength(0);
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

  describe('captureDocState (ADR-0003)', () => {
    it('builds a snapshot whose outline matches the headings, and bumps version each capture', async () => {
      const host = new FakeWordHost();
      host.paragraphs = [
        { text: 'Service Levels', styleBuiltIn: 'Heading1' },
        { text: 'The services are available 99.5% of the time.', styleBuiltIn: 'Normal' },
        { text: 'Availability', styleBuiltIn: 'Heading2' },
      ];
      const bridge = new WordBridge(host);

      const first = await bridge.captureDocState();
      expect(first).toBeDefined();
      if (!first) return;
      expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
      expect(first.surface).toBe('word');
      expect(first.version).toBe(1);
      // The outline comes from the same native heading blocks as resolveContext; the builder
      // strips the Markdown `#` marker so the text is clean (the renderer re-adds it by level).
      expect(first.outline.map((o) => o.text)).toEqual(['Service Levels', 'Availability']);
      expect(first.outline.map((o) => o.level)).toEqual([1, 2]);
      // Headings are content-anchored for re-finding.
      expect(first.outline[0]?.anchor?.matchText).toContain('Service Levels');

      const second = await bridge.captureDocState();
      expect(second?.version).toBe(2);
    });

    it('returns undefined for an empty document', async () => {
      const host = new FakeWordHost();
      host.paragraphs = [];
      expect(await new WordBridge(host).captureDocState()).toBeUndefined();
    });
  });

  describe('searchDocument (ADR-0003 lazy read)', () => {
    it('maps bounded, content-anchored hits to live ResolvedContext', async () => {
      const host = new FakeWordHost();
      host.textHits.set('99.5%', [
        {
          text: 'available 99.5% of the time',
          contextHint: 'Section 5: available 99.5% of the time.',
        },
        { text: 'uptime 99.5%' },
      ]);
      const ctx = await new WordBridge(host).searchDocument('99.5%');

      expect(ctx).toHaveLength(2);
      expect(ctx[0]).toMatchObject({
        ref: { kind: 'selection', surface: 'word', live: true },
        value: { as: 'text' },
      });
      // Content-anchored by the matched text.
      expect(ctx[0]?.ref.anchor?.matchText).toContain('99.5%');
      // The contextHint is folded into the part body so the model sees the surrounding cue.
      if (ctx[0]?.value.as === 'text') expect(ctx[0].value.text).toContain('Section 5');
    });

    it('returns [] for an empty query without touching the host', async () => {
      const host = new FakeWordHost();
      const spy = vi.spyOn(host, 'searchText');
      expect(await new WordBridge(host).searchDocument('  ')).toEqual([]);
      expect(spy).not.toHaveBeenCalled();
    });

    it('returns [] when there are no hits', async () => {
      const host = new FakeWordHost();
      expect(await new WordBridge(host).searchDocument('absent')).toEqual([]);
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
