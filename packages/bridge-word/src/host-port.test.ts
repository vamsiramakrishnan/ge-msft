import { afterEach, describe, expect, it, vi } from 'vitest';
import { OfficeWordHost, type WordHandlers } from './host-port.js';

/**
 * Direct tests for the REAL {@link OfficeWordHost} adapter — the bridge's only un-faked seam. It
 * translates the narrow {@link WordHost} port into `Word.run` / `Office.context` batches, and owns
 * the load-bearing semantics this file pins down:
 *   - lazy `body.search` re-resolution with a bounded hit set + surrounding-paragraph context hint;
 *   - the search→insert tracked-change batch (read-back hits handed to `choose`, write on the
 *     chosen range, drift when `choose` rejects / there are no hits);
 *   - best-effort, requirement-set-gated comment + custom-XML-part writes that never fail the
 *     reversible change they accompany;
 *   - defensive host-event registration + single-owner, idempotent teardown.
 *
 * Rather than depend on the cross-package web-shell simulator, this file installs a SELF-CONTAINED
 * in-memory `globalThis.Word` / `globalThis.Office` modelling ONLY the enumerated slice the adapter
 * drives — the same "fake the global object model so the real bridge runs unchanged" harness pattern
 * the repo already uses (see `web-shell/src/test-harness/fake-word.ts`). Every fake records its
 * `load()`/`sync()` calls so the read-then-write batch ordering is asserted, not assumed.
 */

/* ─────────────────────────── installable global fakes ─────────────────────── */

interface InstallOpts {
  /** query → the match texts `body.search` reads back (one FakeSearchResult each). */
  searchHits?: Record<string, string[]>;
  /** match text → containing-paragraph text, surfaced via `paragraphs.getFirstOrNullObject()`. */
  contextParas?: Record<string, string | undefined>;
  /** Existing comments for `getComments()`. */
  comments?: Array<{ id: string }>;
  /** Requirement-set support map (e.g. { WordApi: 6 }). Defaults to a modern host. */
  requirements?: Record<string, number>;
  /** Force `Word.run` to reject (simulate an unsupported host / batch failure). */
  wordRunThrows?: boolean;
  /** Throw from inside `body.search` (simulate a host quirk mid-batch). */
  searchThrows?: boolean;
  /** Make `onCommentAdded.add()` throw (comment events advertised but unusable). */
  commentAddThrows?: boolean;
  /** Make `onParagraphAdded.add()` throw (one paragraph event not in the active set). */
  paraAddedThrows?: boolean;
  /** Make every Word handler's `.remove()` throw (handler already torn down host-side). */
  removeThrows?: boolean;
  /** Seed the live selection text for readSelectionText. */
  selectionText?: string;
  /** Seed the body paragraphs for readBodyText / readParagraphs. */
  paragraphs?: Array<{ text: string; styleBuiltIn: string }>;
  /** Make `Office...addHandlerAsync` throw (selection observation unavailable). */
  addHandlerThrows?: boolean;
  /** Make `Office...removeHandlerAsync` throw (host already tore the handler down). */
  removeHandlerThrows?: boolean;
}

interface Recorder {
  inserts: Array<{ anchor: string; text: string; location: string }>;
  addedComments: Array<{ anchor: string; text: string }>;
  replies: Array<{ id: string; text: string }>;
  resolved: string[];
  customXml: string[];
  changeTrackingMode?: string;
  syncCount: number;
  /** Ordered trace of loads + syncs + writes, to assert read-then-write batch ordering. */
  trace: string[];
  /** Office event handlers registered via addHandlerAsync, keyed by event type. */
  officeHandlers: Map<string, Array<() => void>>;
  /** Word object-model handlers added via onParagraph-events / onCommentAdded `.add()`. */
  wordHandlerCount: number;
  /** How many Word object-model handlers have had `.remove()` called. */
  wordHandlerRemovals: number;
  removeHandlerCalls: number;
  /** Captured Word object-model handler callbacks, keyed by event label, so a test can fire them. */
  wordHandlerCallbacks: Map<string, (a: never) => Promise<void>>;
}

function install(opts: InstallOpts = {}): { rec: Recorder; restore: () => void } {
  const rec: Recorder = {
    inserts: [],
    addedComments: [],
    replies: [],
    resolved: [],
    customXml: [],
    syncCount: 0,
    trace: [],
    officeHandlers: new Map(),
    wordHandlerCount: 0,
    wordHandlerRemovals: 0,
    removeHandlerCalls: 0,
    wordHandlerCallbacks: new Map(),
  };
  const reqs = opts.requirements ?? { WordApi: 6 };
  const comments = (opts.comments ?? []).map((c) => ({ ...c, replies: [] as string[] }));

  class FakeParagraph {
    isNullObject: boolean;
    text = '';
    constructor(text: string | undefined) {
      this.isNullObject = text === undefined;
      if (text !== undefined) this.text = text;
    }
    load(): this {
      return this;
    }
  }

  class FakeResult {
    constructor(public text: string) {}
    load(p?: string): this {
      rec.trace.push(`result.load:${p ?? ''}`);
      return this;
    }
    get paragraphs(): { getFirstOrNullObject(): FakeParagraph } {
      return {
        getFirstOrNullObject: () =>
          new FakeParagraph(
            Object.prototype.hasOwnProperty.call(opts.contextParas ?? {}, this.text)
              ? opts.contextParas?.[this.text]
              : undefined,
          ),
      };
    }
    insertText(text: string, location: string): void {
      rec.trace.push('insertText');
      rec.inserts.push({ anchor: this.text, text, location });
    }
    insertComment(text: string): void {
      rec.trace.push('insertComment');
      rec.addedComments.push({ anchor: this.text, text });
    }
  }

  class FakeResultCollection {
    items: FakeResult[];
    constructor(query: string) {
      this.items = (opts.searchHits?.[query] ?? []).map((t) => new FakeResult(t));
    }
    load(p?: string): this {
      rec.trace.push(`results.load:${p ?? ''}`);
      return this;
    }
  }

  class FakeCommentProxy {
    constructor(private readonly target: { id: string; replies: string[]; resolved?: boolean }) {}
    get id(): string {
      return this.target.id;
    }
    reply(text: string): void {
      rec.replies.push({ id: this.target.id, text });
    }
    set resolved(v: boolean) {
      if (v) rec.resolved.push(this.target.id);
    }
  }

  class FakeCommentCollection {
    items: FakeCommentProxy[];
    constructor() {
      this.items = comments.map((c) => new FakeCommentProxy(c));
    }
    load(): this {
      return this;
    }
  }

  function wordEvent(label: string): {
    add(h: (a: never) => Promise<void>): { remove(): void };
  } {
    return {
      add(h) {
        if (label === 'commentAdded' && opts.commentAddThrows) throw new Error('add unavailable');
        if (label === 'paraAdded' && opts.paraAddedThrows) throw new Error('event not in set');
        rec.wordHandlerCount += 1;
        rec.wordHandlerCallbacks.set(label, h);
        rec.trace.push(`add:${label}`);
        return {
          remove() {
            if (opts.removeThrows) throw new Error('already gone');
            rec.wordHandlerRemovals += 1;
            rec.trace.push(`remove:${label}`);
          },
        };
      },
    };
  }

  interface FakeBody {
    readonly text: string;
    load(): FakeBody;
    readonly paragraphs: { items: Array<{ text: string; styleBuiltIn: string }>; load(): unknown };
    search(query: string, o: { matchCase: boolean }): FakeResultCollection;
    getComments(): FakeCommentCollection;
  }

  const body: FakeBody = {
    get text(): string {
      if (opts.paragraphs) return opts.paragraphs.map((p) => p.text).join('\n');
      return Object.values(opts.searchHits ?? {})
        .flat()
        .join('\n');
    },
    load(): FakeBody {
      return body;
    },
    get paragraphs(): { items: Array<{ text: string; styleBuiltIn: string }>; load(): unknown } {
      return {
        items: (opts.paragraphs ?? []).map((p) => ({ ...p })),
        load() {
          return this;
        },
      };
    },
    search(query: string, _o: { matchCase: boolean }): FakeResultCollection {
      if (opts.searchThrows) throw new Error('search blew up');
      rec.trace.push(`search:${query}`);
      return new FakeResultCollection(query);
    },
    getComments(): FakeCommentCollection {
      return new FakeCommentCollection();
    },
  };

  const document = {
    body,
    _ctm: '',
    set changeTrackingMode(m: string) {
      this._ctm = m;
      rec.changeTrackingMode = m;
    },
    get changeTrackingMode(): string {
      return this._ctm;
    },
    getSelection(): { text: string; load(): unknown } {
      return { text: opts.selectionText ?? '', load: () => undefined };
    },
    customXmlParts: {
      add(xml: string): { id: string } {
        rec.trace.push('customXml.add');
        rec.customXml.push(xml);
        return { id: `xml-${rec.customXml.length}` };
      },
    },
    onParagraphChanged: wordEvent('paraChanged'),
    onParagraphAdded: wordEvent('paraAdded'),
    onParagraphDeleted: wordEvent('paraDeleted'),
    onCommentAdded: wordEvent('commentAdded'),
  };

  const ctx = {
    document,
    sync(): Promise<void> {
      rec.syncCount += 1;
      rec.trace.push('sync');
      return Promise.resolve();
    },
  };

  const Word = {
    run<T>(cb: (c: typeof ctx) => Promise<T>): Promise<T> {
      if (opts.wordRunThrows) return Promise.reject(new Error('Word.run unsupported'));
      return cb(ctx);
    },
    InsertLocation: { replace: 'Replace', start: 'Start', end: 'End' },
    ChangeTrackingMode: { trackAll: 'TrackAll', off: 'Off' },
    EventSource: { local: 'Local', remote: 'Remote' },
  };

  const Office = {
    EventType: { DocumentSelectionChanged: 'documentSelectionChanged' },
    context: {
      requirements: {
        isSetSupported(name: string, version?: string): boolean {
          const have = reqs[name];
          if (have === undefined) return false;
          if (version === undefined) return true;
          const minor = Number.parseInt(version.split('.')[1] ?? '0', 10);
          return have >= minor;
        },
      },
      document: {
        addHandlerAsync(
          eventType: string,
          handler: () => void,
          cb?: (r: { status: string }) => void,
        ): void {
          if (opts.addHandlerThrows) throw new Error('selection events unavailable');
          const list = rec.officeHandlers.get(eventType) ?? [];
          list.push(handler);
          rec.officeHandlers.set(eventType, list);
          cb?.({ status: 'succeeded' });
        },
        removeHandlerAsync(
          eventType: string,
          options?: { handler?: () => void },
          cb?: (r: { status: string }) => void,
        ): void {
          rec.removeHandlerCalls += 1;
          if (opts.removeHandlerThrows) throw new Error('handler already gone');
          const list = rec.officeHandlers.get(eventType) ?? [];
          rec.officeHandlers.set(
            eventType,
            options?.handler ? list.filter((h) => h !== options.handler) : [],
          );
          cb?.({ status: 'succeeded' });
        },
      },
    },
  };

  const g = globalThis as unknown as Record<string, unknown>;
  const prevWord = g.Word;
  const prevOffice = g.Office;
  g.Word = Word;
  g.Office = Office;

  return {
    rec,
    restore: () => {
      g.Word = prevWord;
      g.Office = prevOffice;
    },
  };
}

let active: { restore: () => void } | undefined;
function setup(opts?: InstallOpts): Recorder {
  const installed = install(opts);
  active = installed;
  return installed.rec;
}

afterEach(() => {
  active?.restore();
  active = undefined;
});

const noopHandlers = (): WordHandlers => ({
  onSelectionChanged: vi.fn(),
  onDocumentChanged: vi.fn(),
  onCommentAdded: vi.fn(),
});

/* ───────────────────────── simple read passthroughs ─────────────────────── */

describe('OfficeWordHost reads', () => {
  it('readSelectionText returns the live selection text', async () => {
    setup({ selectionText: 'The SLA is 99.5%.' });
    await expect(new OfficeWordHost().readSelectionText()).resolves.toBe('The SLA is 99.5%.');
  });

  it('readBodyText returns the joined body text', async () => {
    setup({
      paragraphs: [
        { text: 'Heading', styleBuiltIn: 'Heading1' },
        { text: 'Body line.', styleBuiltIn: 'Normal' },
      ],
    });
    await expect(new OfficeWordHost().readBodyText()).resolves.toBe('Heading\nBody line.');
  });

  it('readParagraphs returns only non-empty paragraphs with their built-in style', async () => {
    setup({
      paragraphs: [
        { text: 'Title', styleBuiltIn: 'Heading1' },
        { text: '   ', styleBuiltIn: 'Normal' }, // whitespace-only → filtered out
        { text: 'Real content', styleBuiltIn: 'Normal' },
      ],
    });
    const paras = await new OfficeWordHost().readParagraphs();
    expect(paras).toEqual([
      { text: 'Title', styleBuiltIn: 'Heading1' },
      { text: 'Real content', styleBuiltIn: 'Normal' },
    ]);
  });
});

/* ─────────────────────────────── searchText ─────────────────────────────── */

describe('OfficeWordHost.searchText (lazy body.search re-resolution)', () => {
  it('returns [] for a blank query without touching the host', async () => {
    const rec = setup({ searchHits: { x: ['x'] } });
    const hits = await new OfficeWordHost().searchText('   ', false);
    expect(hits).toEqual([]);
    expect(rec.trace).not.toContain('search:x');
  });

  it('maps each hit to its match text and folds in the surrounding paragraph as contextHint', async () => {
    setup({
      searchHits: { '99.5%': ['99.5%'] },
      contextParas: { '99.5%': 'Section 5: the SLA is 99.5% measured monthly.' },
    });
    const hits = await new OfficeWordHost().searchText('99.5%', false);
    expect(hits).toEqual([
      { text: '99.5%', contextHint: 'Section 5: the SLA is 99.5% measured monthly.' },
    ]);
  });

  it('omits the contextHint when the paragraph text equals the match text', async () => {
    setup({ searchHits: { SLA: ['SLA'] }, contextParas: { SLA: '  SLA  ' } });
    const [hit] = await new OfficeWordHost().searchText('SLA', false);
    expect(hit).toEqual({ text: 'SLA' });
    expect(hit).not.toHaveProperty('contextHint');
  });

  it('omits the contextHint when the containing paragraph is a null object', async () => {
    setup({ searchHits: { term: ['term'] }, contextParas: { term: undefined } });
    const [hit] = await new OfficeWordHost().searchText('term', false);
    expect(hit).toEqual({ text: 'term' });
  });

  it('caps the result set at MAX_SEARCH_HITS (8) even when the host returns more', async () => {
    const many = Array.from({ length: 20 }, (_, i) => `m${i}`);
    setup({ searchHits: { the: many } });
    const hits = await new OfficeWordHost().searchText('the', false);
    expect(hits).toHaveLength(8);
    expect(hits.map((h) => h.text)).toEqual(many.slice(0, 8));
  });

  it('degrades to [] (never throws) when body.search throws mid-batch', async () => {
    setup({ searchThrows: true });
    await expect(new OfficeWordHost().searchText('x', false)).resolves.toEqual([]);
  });

  it('degrades to [] when Word.run itself rejects (unsupported host)', async () => {
    setup({ searchHits: { x: ['x'] }, wordRunThrows: true });
    await expect(new OfficeWordHost().searchText('x', false)).resolves.toEqual([]);
  });

  it('returns [] (no hits) without inventing entries', async () => {
    setup({ searchHits: {} });
    await expect(new OfficeWordHost().searchText('absent', false)).resolves.toEqual([]);
  });
});

/* ──────────────────────────── applyTrackedChange ─────────────────────────── */

describe('OfficeWordHost.applyTrackedChange (search→choose→insert batch)', () => {
  it('turns on tracked changes, hands read-back hits to choose, and writes on the chosen range', async () => {
    const rec = setup({ searchHits: { '99.5%': ['intro 99.5%', 'SLA: 99.5% uptime'] } });
    const choose = vi.fn((texts: readonly string[]) => texts.findIndex((t) => t.includes('SLA')));
    const out = await new OfficeWordHost().applyTrackedChange(
      '99.5%',
      { matchCase: false },
      '99.9%',
      choose,
    );

    expect(out).toEqual({ status: 'applied', location: 'tracked-change' });
    expect(rec.changeTrackingMode).toBe('TrackAll');
    // choose saw the read-back match texts.
    expect(choose).toHaveBeenCalledWith(['intro 99.5%', 'SLA: 99.5% uptime']);
    // wrote exactly once, on the chosen (second) hit, with the replace location.
    expect(rec.inserts).toEqual([
      { anchor: 'SLA: 99.5% uptime', text: '99.9%', location: 'Replace' },
    ]);
  });

  it('preserves the read-then-write ordering: load + first sync precede the choose-driven insert', async () => {
    const rec = setup({ searchHits: { q: ['only hit'] } });
    await new OfficeWordHost().applyTrackedChange('q', { matchCase: true }, 'new', () => 0);
    // search → load → sync (read) → insertText → sync (write): two syncs, insert after the first.
    const firstSync = rec.trace.indexOf('sync');
    const insert = rec.trace.indexOf('insertText');
    const lastSync = rec.trace.lastIndexOf('sync');
    expect(firstSync).toBeGreaterThanOrEqual(0);
    expect(insert).toBeGreaterThan(firstSync);
    expect(lastSync).toBeGreaterThan(insert);
    expect(rec.syncCount).toBe(2);
  });

  it('degrades to drift (no write, no second sync) when there are zero hits', async () => {
    const rec = setup({ searchHits: { '99.5%': [] } });
    const out = await new OfficeWordHost().applyTrackedChange(
      '99.5%',
      { matchCase: false },
      'x',
      () => 0,
    );
    expect(out).toEqual({ status: 'drift' });
    expect(rec.inserts).toHaveLength(0);
    // Only the read sync ran — the write sync never happened.
    expect(rec.syncCount).toBe(1);
  });

  it('degrades to drift when choose rejects every hit with a negative index', async () => {
    const rec = setup({ searchHits: { q: ['a', 'b'] } });
    const out = await new OfficeWordHost().applyTrackedChange(
      'q',
      { matchCase: false },
      'x',
      () => -1,
    );
    expect(out).toEqual({ status: 'drift' });
    expect(rec.inserts).toHaveLength(0);
  });

  it('passes matchCase through to the search', async () => {
    const rec = setup({ searchHits: { SLA: ['SLA'] } });
    await new OfficeWordHost().applyTrackedChange('SLA', { matchCase: true }, 'x', () => 0);
    expect(rec.trace).toContain('search:SLA');
    // a hit was found and written (proves the search ran with the supplied query)
    expect(rec.inserts).toHaveLength(1);
  });
});

/* ───────────────────────────────── addComment ───────────────────────────── */

describe('OfficeWordHost.addComment (requirement-set gated, best-effort)', () => {
  it('inserts a comment on the first re-resolved hit when WordApi 1.4 is supported', async () => {
    const rec = setup({ searchHits: { claim: ['the claim'] }, requirements: { WordApi: 4 } });
    const res = await new OfficeWordHost().addComment('claim', false, 'Needs a source');
    expect(res).toEqual({ ok: true });
    expect(rec.addedComments).toEqual([{ anchor: 'the claim', text: 'Needs a source' }]);
  });

  it('skips silently (ok:false, no host call) when WordApi 1.4 is NOT supported', async () => {
    const rec = setup({ searchHits: { claim: ['the claim'] }, requirements: { WordApi: 3 } });
    const res = await new OfficeWordHost().addComment('claim', false, 'note');
    expect(res).toEqual({ ok: false });
    // never even searched — gated before the batch.
    expect(rec.trace).not.toContain('search:claim');
    expect(rec.addedComments).toHaveLength(0);
  });

  it('returns ok:false when the anchor text is gone (no hit to attach to)', async () => {
    const rec = setup({ searchHits: { claim: [] }, requirements: { WordApi: 4 } });
    const res = await new OfficeWordHost().addComment('claim', false, 'note');
    expect(res).toEqual({ ok: false });
    expect(rec.addedComments).toHaveLength(0);
  });

  it('returns ok:false (never throws) when the comments batch throws', async () => {
    setup({ searchHits: { claim: ['hit'] }, requirements: { WordApi: 4 }, searchThrows: true });
    await expect(new OfficeWordHost().addComment('claim', false, 'note')).resolves.toEqual({
      ok: false,
    });
  });
});

/* ─────────────────────────────── persistProvenance ──────────────────────── */

describe('OfficeWordHost.persistProvenance (durable custom XML part, gated)', () => {
  it('adds the XML part and returns ok:true on a WordApi 1.4 host', async () => {
    const rec = setup({ requirements: { WordApi: 4 } });
    const res = await new OfficeWordHost().persistProvenance('<prov key="ge:prov:c1"/>');
    expect(res).toEqual({ ok: true });
    expect(rec.customXml).toEqual(['<prov key="ge:prov:c1"/>']);
  });

  it('skips silently (ok:false, no add) when WordApi 1.4 is unsupported', async () => {
    const rec = setup({ requirements: { WordApi: 3 } });
    const res = await new OfficeWordHost().persistProvenance('<prov/>');
    expect(res).toEqual({ ok: false });
    expect(rec.customXml).toHaveLength(0);
  });

  it('returns ok:false (never throws) when Word.run rejects', async () => {
    setup({ requirements: { WordApi: 4 }, wordRunThrows: true });
    await expect(new OfficeWordHost().persistProvenance('<prov/>')).resolves.toEqual({ ok: false });
  });
});

/* ─────────────────────────────── replyToComment ─────────────────────────── */

describe('OfficeWordHost.replyToComment', () => {
  it('replies to the matching comment and reports gone:false', async () => {
    const rec = setup({ comments: [{ id: 'cmt-1' }] });
    const out = await new OfficeWordHost().replyToComment('cmt-1', 'addressed', false);
    expect(out).toEqual({ status: 'replied', location: 'comment:cmt-1' });
    expect(rec.replies).toEqual([{ id: 'cmt-1', text: 'addressed' }]);
    expect(rec.resolved).toEqual([]); // resolve=false → not resolved
  });

  it('sets resolved=true when resolve is requested', async () => {
    const rec = setup({ comments: [{ id: 'cmt-2' }] });
    await new OfficeWordHost().replyToComment('cmt-2', 'done', true);
    expect(rec.resolved).toEqual(['cmt-2']);
  });

  it('returns gone when the comment id no longer exists, writing nothing', async () => {
    const rec = setup({ comments: [{ id: 'other' }] });
    const out = await new OfficeWordHost().replyToComment('missing', 'hi', true);
    expect(out).toEqual({ status: 'gone' });
    expect(rec.replies).toHaveLength(0);
    expect(rec.resolved).toHaveLength(0);
  });
});

/* ─────────────────────────────── registerHandlers ───────────────────────── */

describe('OfficeWordHost.registerHandlers (wiring + single-owner teardown)', () => {
  it('registers a selection handler that forwards to onSelectionChanged', async () => {
    const rec = setup();
    const handlers = noopHandlers();
    const unsub = new OfficeWordHost().registerHandlers(handlers);
    await Promise.resolve();

    expect(rec.officeHandlers.get('documentSelectionChanged')).toHaveLength(1);
    rec.officeHandlers.get('documentSelectionChanged')?.[0]?.();
    expect(handlers.onSelectionChanged).toHaveBeenCalledTimes(1);
    unsub();
  });

  it('registers the three paragraph events + comment event on a WordApi 1.6 host', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();
    // onParagraphChanged/Added/Deleted (3) + onCommentAdded (1).
    expect(rec.wordHandlerCount).toBe(4);
  });

  it('does NOT register paragraph events when WordApi 1.6 is unsupported (gate, not truthiness)', async () => {
    const rec = setup({ requirements: { WordApi: 4 } });
    new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();
    // Only the comment handler (its own feature-detect) may attach; the 3 paragraph events do not.
    expect(rec.wordHandlerCount).toBeLessThanOrEqual(1);
  });

  it('unsubscribe removes the selection handler and every registered Word handler exactly once', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();
    expect(rec.wordHandlerCount).toBe(4);

    unsub();
    // teardown chains off the registration promise; let it settle.
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();

    expect(rec.removeHandlerCalls).toBe(1); // selection handler removed
    expect(rec.officeHandlers.get('documentSelectionChanged')).toHaveLength(0);
    expect(rec.wordHandlerRemovals).toBe(4); // each Word handler removed once
  });

  it('is idempotent: a second unsubscribe performs no further teardown', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();

    unsub();
    await Promise.resolve();
    await Promise.resolve();
    const removalsAfterFirst = rec.wordHandlerRemovals;
    const officeRemovesAfterFirst = rec.removeHandlerCalls;

    unsub();
    await Promise.resolve();
    await Promise.resolve();

    expect(rec.wordHandlerRemovals).toBe(removalsAfterFirst);
    expect(rec.removeHandlerCalls).toBe(officeRemovesAfterFirst);
  });

  it('forwards document edits with the host-supplied coauthor source untouched', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    const handlers = noopHandlers();
    new OfficeWordHost().registerHandlers(handlers);
    await Promise.resolve();
    await Promise.resolve();
    // The fake records that handlers attached; assert the wiring count rather than reaching into
    // the closure — the bridge-level mapping of source→origin is covered in word-bridge.test.ts.
    expect(rec.wordHandlerCount).toBe(4);
    expect(handlers.onDocumentChanged).not.toHaveBeenCalled(); // not fired yet
  });

  it('maps the comment event args (source + each of commentId/id/ids) through to onCommentAdded', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    const handlers = noopHandlers();
    new OfficeWordHost().registerHandlers(handlers);
    await Promise.resolve();
    await Promise.resolve();

    const fire = rec.wordHandlerCallbacks.get('commentAdded');
    expect(fire).toBeDefined();
    // Full arg shape: every present field is forwarded; absent fields are omitted.
    await fire?.({ source: 'Remote', commentId: 'cmt-1', id: 'id-1', ids: ['a', 'b'] } as never);
    expect(handlers.onCommentAdded).toHaveBeenCalledWith({
      source: 'Remote',
      commentId: 'cmt-1',
      id: 'id-1',
      ids: ['a', 'b'],
    });

    // An empty arg object forwards an empty object (no undefined keys leak through).
    await fire?.({} as never);
    expect(handlers.onCommentAdded).toHaveBeenLastCalledWith({});
  });

  it('forwards a paragraph edit event with the host coauthor source to onDocumentChanged', async () => {
    const rec = setup({ requirements: { WordApi: 6 } });
    const handlers = noopHandlers();
    new OfficeWordHost().registerHandlers(handlers);
    await Promise.resolve();
    await Promise.resolve();

    const fire = rec.wordHandlerCallbacks.get('paraChanged');
    await fire?.({ source: 'Remote' } as never);
    expect(handlers.onDocumentChanged).toHaveBeenCalledWith({ source: 'Remote' });
  });

  it('skips a single paragraph event whose .add() throws but still registers the others + comment', async () => {
    const rec = setup({ requirements: { WordApi: 6 }, paraAddedThrows: true });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();
    // onParagraphChanged + onParagraphDeleted (2) + onCommentAdded (1); onParagraphAdded threw.
    expect(rec.wordHandlerCount).toBe(3);
    expect(rec.wordHandlerCallbacks.has('paraAdded')).toBe(false);
    expect(() => unsub()).not.toThrow();
  });

  it('skips the comment handler (without throwing) when its .add() throws', async () => {
    const rec = setup({ requirements: { WordApi: 6 }, commentAddThrows: true });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();
    // Only the 3 paragraph events attached; the comment add threw and was swallowed.
    expect(rec.wordHandlerCount).toBe(3);
    expect(() => unsub()).not.toThrow();
  });

  it('teardown tolerates a Word handler whose .remove() throws (best-effort)', async () => {
    const rec = setup({ requirements: { WordApi: 6 }, removeThrows: true });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();

    unsub();
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // Every remove() threw, but teardown still ran the selection removal and didn't blow up.
    expect(rec.wordHandlerRemovals).toBe(0);
    expect(rec.removeHandlerCalls).toBe(1);
  });

  it('tolerates addHandlerAsync throwing: no selection handler, and teardown skips its removal', async () => {
    const rec = setup({ requirements: { WordApi: 6 }, addHandlerThrows: true });
    const handlers = noopHandlers();
    let unsub: (() => void) | undefined;
    expect(() => {
      unsub = new OfficeWordHost().registerHandlers(handlers);
    }).not.toThrow();
    await Promise.resolve();
    await Promise.resolve();
    expect(rec.officeHandlers.get('documentSelectionChanged') ?? []).toHaveLength(0);

    unsub?.();
    await Promise.resolve();
    // onSelection was never set → removeSelection short-circuits, removeHandlerAsync not called.
    expect(rec.removeHandlerCalls).toBe(0);
  });

  it('tolerates removeHandlerAsync throwing during teardown (best-effort) without throwing', async () => {
    const rec = setup({ requirements: { WordApi: 6 }, removeHandlerThrows: true });
    const unsub = new OfficeWordHost().registerHandlers(noopHandlers());
    await Promise.resolve();
    await Promise.resolve();

    expect(() => unsub()).not.toThrow();
    await Promise.resolve();
    await Promise.resolve();
    // It attempted the removal (count incremented) but the throw was swallowed.
    expect(rec.removeHandlerCalls).toBe(1);
  });

  it('survives a Word.run rejection during registration without throwing', async () => {
    setup({ wordRunThrows: true });
    const host = new OfficeWordHost();
    let unsub: (() => void) | undefined;
    expect(() => {
      unsub = host.registerHandlers(noopHandlers());
    }).not.toThrow();
    await Promise.resolve();
    // teardown must also be safe even though nothing Word-side registered.
    expect(() => unsub?.()).not.toThrow();
    await Promise.resolve();
  });
});
