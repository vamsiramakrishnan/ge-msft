import { afterEach, describe, it, expect } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationRequest,
  type ContextRef,
} from '@ge/contracts';
import { HANDLED_ACTUATIONS, OneNoteBridge } from './onenote-bridge.js';

/**
 * Behavioural tests for the {@link OneNoteBridge} host wiring — the file that drives the
 * `OneNote.run` / `Office` globals directly. There is no host-port indirection here, so we install
 * a self-contained in-memory OneNote + Office simulator onto `globalThis` for each test. The
 * simulator re-implements ONLY the slice of the `OneNoteApi` object model the bridge actually
 * drives: active page → contents → outline → paragraphs → richText (the READ paths), and active
 * section → addPage → addOutline (the WRITE / page-synthesis path), plus the requirement gate. The
 * REAL bridge runs unchanged against it, so these assert real semantics: capability gating, the
 * batched read, empty-page degradation, search, the append-page synthesis + citation-tagged outline
 * insertion, the no-section / unsupported-host degradations, and changeId echo.
 */

/* ─────────────────────────── in-memory host model ─────────────────────────── */

interface ParaSeed {
  /** 'RichText' (carries text) or any other type (skipped by the bridge). */
  type: string;
  text?: string;
}
interface ContentSeed {
  /** 'Outline' (has paragraphs) or any other type (skipped). */
  type: string;
  paragraphs?: ParaSeed[];
}
interface PageSeed {
  id: string;
  title: string;
  contents: ContentSeed[];
}
interface AddedPageRecord {
  id: string;
  title: string;
  outlineHtml: string | undefined;
  outlineLeft: number | undefined;
  outlineTop: number | undefined;
}
interface SectionSeed {
  id: string;
  /** Pages added during a write, in order, with the outline HTML they received. */
  addedPages: AddedPageRecord[];
}
interface NotebookSeed {
  /** The active page, or null to simulate "no active page". */
  page: PageSeed | null;
  /** The active section, or null to simulate "no active section". */
  section: SectionSeed | null;
  navigatedPages: string[];
  navigatedUrls: string[];
}

class FakeRichText {
  constructor(private readonly seed: ParaSeed) {}
  text: string | undefined = undefined;
  load(_p?: string): this {
    // The bridge reads `.text` only after a sync; mirror the seed at load time.
    this.text = this.seed.text;
    return this;
  }
}

class FakeParagraph {
  readonly richText: FakeRichText;
  constructor(private readonly seed: ParaSeed) {
    this.richText = new FakeRichText(seed);
  }
  get type(): string {
    return this.seed.type;
  }
}

class FakeParagraphCollection {
  items: FakeParagraph[] = [];
  constructor(private readonly seed: ParaSeed[]) {}
  load(_p?: string): this {
    this.items = this.seed.map((p) => new FakeParagraph(p));
    return this;
  }
}

class FakeOutline {
  readonly paragraphs: FakeParagraphCollection;
  constructor(seed: ParaSeed[]) {
    this.paragraphs = new FakeParagraphCollection(seed);
  }
}

class FakePageContent {
  readonly outline: FakeOutline;
  constructor(private readonly seed: ContentSeed) {
    this.outline = new FakeOutline(seed.paragraphs ?? []);
  }
  get type(): string {
    return this.seed.type;
  }
}

class FakePageContentCollection {
  items: FakePageContent[] = [];
  constructor(private readonly seed: ContentSeed[]) {}
  load(_p?: string): this {
    this.items = this.seed.map((c) => new FakePageContent(c));
    return this;
  }
}

class FakePage {
  readonly isNullObject: boolean;
  readonly contents: FakePageContentCollection;
  id = '';
  title = '';
  /** addOutline calls recorded for the write path. */
  outlineCalls: Array<{ left: number; top: number; html: string }> = [];
  constructor(private readonly seed: PageSeed | null) {
    this.isNullObject = seed === null;
    this.contents = new FakePageContentCollection(seed?.contents ?? []);
  }
  load(props?: string): this {
    if (this.seed) {
      if (props?.includes('id')) this.id = this.seed.id;
      if (props?.includes('title')) this.title = this.seed.title;
    }
    return this;
  }
  addOutline(left: number, top: number, html: string): void {
    this.outlineCalls.push({ left, top, html });
  }
}

class FakeSection {
  readonly isNullObject: boolean;
  id = '';
  /** The page returned by addPage so tests can inspect its outline. */
  lastAddedPage: FakePage | undefined;
  private counter = 0;
  constructor(private readonly seed: SectionSeed | null) {
    this.isNullObject = seed === null;
  }
  load(props?: string): this {
    if (this.seed && props?.includes('id')) this.id = this.seed.id;
    return this;
  }
  addPage(title: string): FakePage {
    if (!this.seed) throw new Error('fake-onenote: addPage on null section');
    this.counter += 1;
    const id = `${this.seed.id}-page-${this.counter}`;
    const pageSeed: PageSeed = { id, title, contents: [] };
    const page = new FakePage(pageSeed);
    const record: AddedPageRecord = {
      id,
      title,
      outlineHtml: undefined,
      outlineLeft: undefined,
      outlineTop: undefined,
    };
    this.seed.addedPages.push(record);
    // Mirror outline writes back into the record for assertions.
    const origAddOutline = page.addOutline.bind(page);
    page.addOutline = (l: number, t: number, html: string): void => {
      record.outlineHtml = html;
      record.outlineLeft = l;
      record.outlineTop = t;
      origAddOutline(l, t, html);
    };
    this.lastAddedPage = page;
    return page;
  }
}

class FakeApplication {
  readonly page: FakePage;
  readonly section: FakeSection;
  pageRequested = false;
  sectionRequested = false;
  constructor(private readonly seed: NotebookSeed) {
    this.page = new FakePage(seed.page);
    this.section = new FakeSection(seed.section);
  }
  getActivePageOrNull(): FakePage {
    this.pageRequested = true;
    return this.page;
  }
  getActiveSectionOrNull(): FakeSection {
    this.sectionRequested = true;
    return this.section;
  }
  navigateToPage(page: FakePage): void {
    this.seed.navigatedPages.push(page.id);
  }
  navigateToPageWithClientUrl(url: string): FakePage {
    this.seed.navigatedUrls.push(url);
    return this.page;
  }
}

class FakeContext {
  readonly application: FakeApplication;
  syncs = 0;
  constructor(seed: NotebookSeed) {
    this.application = new FakeApplication(seed);
  }
  sync(): Promise<void> {
    this.syncs += 1;
    return Promise.resolve();
  }
}

interface InstallOpts {
  /** Requirement-set support: name → max supported version (numeric compare). */
  requirements?: Record<string, number>;
}

interface Installed {
  ctxRef: { ctx?: FakeContext };
  restore(): void;
}

function install(seed: NotebookSeed, opts: InstallOpts = {}): Installed {
  const reqs = opts.requirements ?? { OneNoteApi: 1.1 };
  const ctxRef: { ctx?: FakeContext } = {};

  const prevOneNote = (globalThis as Record<string, unknown>).OneNote;
  const prevOffice = (globalThis as Record<string, unknown>).Office;

  (globalThis as Record<string, unknown>).OneNote = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeContext(seed);
      ctxRef.ctx = ctx;
      return cb(ctx);
    },
  };

  (globalThis as Record<string, unknown>).Office = {
    context: {
      requirements: {
        isSetSupported: (name: string, version?: string): boolean => {
          const max = reqs[name];
          if (max === undefined) return false;
          const v = version ? Number.parseFloat(version) : 0;
          return v <= max;
        },
      },
    },
  };

  return {
    ctxRef,
    restore: () => {
      (globalThis as Record<string, unknown>).OneNote = prevOneNote;
      (globalThis as Record<string, unknown>).Office = prevOffice;
    },
  };
}

let installed: Installed | undefined;
afterEach(() => {
  installed?.restore();
  installed = undefined;
});

/* ─────────────────────────── fixtures ─────────────────────────── */

function outline(paragraphs: ParaSeed[]): ContentSeed {
  return { type: 'Outline', paragraphs };
}

const SAMPLE_PAGE: PageSeed = {
  id: 'pg-1',
  title: 'Source review',
  contents: [
    outline([
      { type: 'RichText', text: 'Northwind MSA v3 is current.' },
      { type: 'RichText', text: '   ' }, // blank → dropped
      { type: 'RichText', text: 'ISO 27001 valid through Nov 2026.' },
    ]),
  ],
};

function notebook(page: PageSeed | null, section: SectionSeed | null = null): NotebookSeed {
  // Deep-copy the page so write paths can't mutate the shared fixture across tests.
  const pageCopy: PageSeed | null = page
    ? {
        id: page.id,
        title: page.title,
        contents: page.contents.map((c) => ({
          type: c.type,
          paragraphs: c.paragraphs?.map((p) => ({ ...p })),
        })),
      }
    : null;
  return { page: pageCopy, section, navigatedPages: [], navigatedUrls: [] };
}

function section(id = 'sec-1'): SectionSeed {
  return { id, addedPages: [] };
}

function appendPage(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'append-page', surface: 'onenote', params };
}

/* ───────────────────────────── surface + capabilities ───────────────────────────── */

describe('OneNoteBridge surface + capabilities', () => {
  it('reports the onenote surface and a manifest advertising append-page', () => {
    const bridge = new OneNoteBridge();
    expect(bridge.surface).toBe('onenote');
    const caps = bridge.getCapabilities();
    expect(caps.surface).toBe('onenote');
    expect(caps.actuations.map((a) => a.kind)).toEqual(['append-page']);
    expect(HANDLED_ACTUATIONS).toEqual(['append-page']);
  });
});

/* ───────────────────────────── listContext ───────────────────────────── */

describe('OneNoteBridge.listContext', () => {
  it('returns [] without touching the host when OneNoteApi 1.1 is unsupported', async () => {
    installed = install(notebook(SAMPLE_PAGE), { requirements: { OneNoteApi: 1.0 } });
    expect(await new OneNoteBridge().listContext()).toEqual([]);
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('lists the active page as a live page ref anchored by page id', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    const refs = await new OneNoteBridge().listContext();
    expect(refs).toEqual([
      {
        id: 'on:page:pg-1',
        kind: 'page',
        surface: 'onenote',
        title: 'Source review',
        live: true,
      },
    ]);
  });

  it('falls back to a "Current page" title when the page has no title', async () => {
    installed = install(notebook({ id: 'pg-9', title: '', contents: [] }));
    const refs = await new OneNoteBridge().listContext();
    expect(refs[0]?.title).toBe('Current page');
  });

  it('returns [] when there is no active page (null object)', async () => {
    installed = install(notebook(null));
    expect(await new OneNoteBridge().listContext()).toEqual([]);
  });
});

/* ───────────────────────────── resolveContext ───────────────────────────── */

describe('OneNoteBridge.resolveContext', () => {
  const ref: ContextRef = {
    id: 'on:page:pg-1',
    kind: 'page',
    surface: 'onenote',
    title: 'Source review',
  };

  it('resolves the active page to valid context anchored by page id, dropping blank paragraphs', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    const ctx = await new OneNoteBridge().resolveContext(ref);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'page:pg-1')).toBe(true);
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('Northwind');
    expect(joined).toContain('ISO 27001');
  });

  it('returns [] when no active page exists', async () => {
    installed = install(notebook(null));
    expect(await new OneNoteBridge().resolveContext(ref)).toEqual([]);
  });

  it('returns [] on an older host (OneNoteApi 1.1 unsupported)', async () => {
    installed = install(notebook(SAMPLE_PAGE), { requirements: { OneNoteApi: 1.0 } });
    expect(await new OneNoteBridge().resolveContext(ref)).toEqual([]);
  });

  it('only reads RichText paragraphs from Outline contents, skipping other types', async () => {
    const mixed: PageSeed = {
      id: 'pg-m',
      title: 'Mixed',
      contents: [
        { type: 'Image' }, // non-outline content → skipped
        outline([
          { type: 'RichText', text: 'kept paragraph' },
          { type: 'Table' }, // non-richtext paragraph → skipped, no text
        ]),
      ],
    };
    installed = install(notebook(mixed));
    const ctx = await new OneNoteBridge().resolveContext(ref);
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('kept paragraph');
  });
});

/* ───────────────────────────── revealContext ───────────────────────────── */

describe('OneNoteBridge.revealContext', () => {
  it('navigates to the active page when the page ref matches', async () => {
    const seed = notebook(SAMPLE_PAGE);
    installed = install(seed);
    const bridge = new OneNoteBridge();
    const ref: ContextRef = {
      id: 'on:page:pg-1',
      kind: 'page',
      surface: 'onenote',
      title: 'Source review',
    };

    expect(bridge.canRevealContext(ref)).toBe(true);
    await bridge.revealContext(ref);

    expect(seed.navigatedPages).toEqual(['pg-1']);
  });

  it('can navigate a page client URL when supplied as an anchor locator', async () => {
    const seed = notebook(SAMPLE_PAGE);
    installed = install(seed);
    await new OneNoteBridge().revealContext({
      id: 'on:page:external',
      kind: 'page',
      surface: 'onenote',
      title: 'External page',
      anchor: {
        matchText: 'External page',
        locator: 'clientUrl:https://contoso.example/onenote/page',
      },
    });

    expect(seed.navigatedUrls).toEqual(['https://contoso.example/onenote/page']);
  });

  it('does not navigate when the active page differs from the requested page id', async () => {
    const seed = notebook(SAMPLE_PAGE);
    installed = install(seed);
    await new OneNoteBridge().revealContext({
      id: 'on:page:other',
      kind: 'page',
      surface: 'onenote',
      title: 'Other page',
    });

    expect(seed.navigatedPages).toEqual([]);
  });
});

/* ───────────────────────────── captureDocState ───────────────────────────── */

describe('OneNoteBridge.captureDocState', () => {
  it('returns undefined on an older host', async () => {
    installed = install(notebook(SAMPLE_PAGE), { requirements: { OneNoteApi: 1.0 } });
    expect(await new OneNoteBridge().captureDocState()).toBeUndefined();
  });

  it('returns undefined when there is no active page', async () => {
    installed = install(notebook(null));
    expect(await new OneNoteBridge().captureDocState()).toBeUndefined();
  });

  it('returns undefined for a page with no usable blocks (empty, untitled)', async () => {
    installed = install(notebook({ id: 'pg-e', title: '   ', contents: [outline([])] }));
    expect(await new OneNoteBridge().captureDocState()).toBeUndefined();
  });

  it('builds a valid snapshot from the active page and bumps the version each capture', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    const bridge = new OneNoteBridge();

    const first = await bridge.captureDocState();
    expect(first).toBeDefined();
    if (!first) return;
    expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
    expect(first.surface).toBe('onenote');
    expect(first.version).toBe(1);
    expect(first.outline.some((o) => o.text.includes('Source review'))).toBe(true);

    const second = await bridge.captureDocState();
    expect(second?.version).toBe(2);
  });

  it('omits the title from the snapshot when the page title is only whitespace', async () => {
    // A whitespace title with at least one real paragraph still yields blocks (paragraph),
    // but the snapshot title must be omitted (the `page.title.trim()` guard).
    const page: PageSeed = {
      id: 'pg-w',
      title: '   ',
      contents: [outline([{ type: 'RichText', text: 'a real paragraph' }])],
    };
    installed = install(notebook(page));
    const snap = await new OneNoteBridge().captureDocState();
    expect(snap).toBeDefined();
    if (!snap) return;
    expect(snap.title).toBeUndefined();
  });
});

/* ───────────────────────────── searchDocument ───────────────────────────── */

describe('OneNoteBridge.searchDocument', () => {
  it('returns [] for an empty / whitespace query without touching the host', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    expect(await new OneNoteBridge().searchDocument('   ')).toEqual([]);
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('returns [] on an older host even for a real query', async () => {
    installed = install(notebook(SAMPLE_PAGE), { requirements: { OneNoteApi: 1.0 } });
    expect(await new OneNoteBridge().searchDocument('ISO')).toEqual([]);
  });

  it('returns [] when there is no active page', async () => {
    installed = install(notebook(null));
    expect(await new OneNoteBridge().searchDocument('ISO')).toEqual([]);
  });

  it('returns the matching paragraphs (case-insensitive), excluding non-matches', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    const ctx = await new OneNoteBridge().searchDocument('iso');
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx
      .map((c) => (c.value.as === 'text' ? c.value.text : ''))
      .join('\n')
      .toLowerCase();
    expect(joined).toContain('iso 27001');
    expect(joined).not.toContain('northwind');
  });

  it('returns [] when nothing matches the query', async () => {
    installed = install(notebook(SAMPLE_PAGE));
    expect(await new OneNoteBridge().searchDocument('nonexistent-token')).toEqual([]);
  });
});

/* ───────────────────────────── actuate: dispatch ───────────────────────────── */

describe('OneNoteBridge.actuate dispatch', () => {
  it('rejects a non-append-page kind as unsupported, echoing the changeId, without touching the host', async () => {
    installed = install(notebook(SAMPLE_PAGE, section()));
    const res = await new OneNoteBridge().actuate({
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      surface: 'excel',
      params: { cells: [['1']] },
    });
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      error: { code: 'unsupported' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('rejects an empty synthesis (no html/text) before touching the host', async () => {
    installed = install(notebook(SAMPLE_PAGE, section()));
    const res = await new OneNoteBridge().actuate(appendPage({}, 'chg-empty'));
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-empty'),
      kind: 'append-page',
      error: { code: 'empty_synthesis' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('degrades when the OneNote write API is unsupported (older host)', async () => {
    installed = install(notebook(SAMPLE_PAGE, section()), { requirements: { OneNoteApi: 1.0 } });
    const res = await new OneNoteBridge().actuate(
      appendPage({ text: 'A grounded claim.' }, 'chg-old'),
    );
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-old'),
      kind: 'append-page',
      degraded: true,
      error: { code: 'onenote_unsupported' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('degrades when there is no active section to add the page to', async () => {
    installed = install(notebook(SAMPLE_PAGE, null));
    const res = await new OneNoteBridge().actuate(
      appendPage({ text: 'A grounded claim.' }, 'chg-nosec'),
    );
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-nosec'),
      kind: 'append-page',
      degraded: true,
      error: { code: 'no_section' },
    });
  });
});

/* ───────────────────────────── actuate: append-page write ───────────────────────────── */

describe('OneNoteBridge.actuate append-page (write path)', () => {
  it('adds a page titled from the target matchText and writes a citation-tagged outline', async () => {
    const sec = section();
    installed = install(notebook(SAMPLE_PAGE, sec));
    const res = await new OneNoteBridge().actuate(
      appendPage(
        {
          text: 'The SLA sits below standard.',
          sources: [{ title: 'Risk Policy', locator: '§3.2' }],
          target: { matchText: 'Risk synthesis' },
        },
        'chg-write',
      ),
    );
    expect(res).toEqual({
      ok: true,
      changeId: asChangeId('chg-write'),
      kind: 'append-page',
      location: `page:${sec.addedPages[0]?.id}`,
    });
    // Exactly one page added, titled from matchText.
    expect(sec.addedPages).toHaveLength(1);
    expect(sec.addedPages[0]?.title).toBe('Risk synthesis');
    // The outline carries the synthesized claim AND an inline citation tag (provenance).
    const html = sec.addedPages[0]?.outlineHtml ?? '';
    expect(html).toContain('The SLA sits below standard.');
    expect(html).toContain('data-ge-cite="1"');
    expect(html).toContain('[Risk Policy · §3.2]');
  });

  it('defaults the page title to "Synthesis" when no matchText is supplied', async () => {
    const sec = section();
    installed = install(notebook(SAMPLE_PAGE, sec));
    await new OneNoteBridge().actuate(appendPage({ text: 'A claim.' }));
    expect(sec.addedPages[0]?.title).toBe('Synthesis');
  });

  it('uses prebuilt html verbatim (already citation-tagged) for the outline', async () => {
    const sec = section();
    installed = install(notebook(SAMPLE_PAGE, sec));
    const prebuilt = '<p>Already tagged <span data-ge-cite="1">[x]</span></p>';
    const res = await new OneNoteBridge().actuate(appendPage({ html: prebuilt }));
    expect(res.ok).toBe(true);
    expect(sec.addedPages[0]?.outlineHtml).toBe(prebuilt);
  });

  it('treats untrusted plain-text synthesis as data — HTML in the text is escaped, not injected', async () => {
    const sec = section();
    installed = install(notebook(SAMPLE_PAGE, sec));
    await new OneNoteBridge().actuate(appendPage({ text: '<img src=x onerror=alert(1)>' }));
    const html = sec.addedPages[0]?.outlineHtml ?? '';
    expect(html).toContain('&lt;img src=x onerror=alert(1)&gt;');
    expect(html).not.toContain('<img');
  });

  it('positions the outline at the standard top-left offset (40,40)', async () => {
    const sec = section();
    installed = install(notebook(SAMPLE_PAGE, sec));
    await new OneNoteBridge().actuate(appendPage({ text: 'A claim.' }));
    expect(sec.addedPages).toHaveLength(1);
    expect(sec.addedPages[0]?.outlineLeft).toBe(40);
    expect(sec.addedPages[0]?.outlineTop).toBe(40);
  });
});
