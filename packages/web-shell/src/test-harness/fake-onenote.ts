/**
 * In-memory **OneNote host simulator**. Models the slice of the `OneNoteApi` (legacy XML add-in)
 * object model the real {@link "@ge/bridge-onenote"!OneNoteBridge} drives — its only un-faked seam —
 * so the REAL bridge runs unchanged against a seeded notebook through `OneNote.run`.
 *
 * The OneNote global name is NOT one of the harness's `installGlobal` host names, and this fake is
 * imported by RELATIVE path (it is intentionally OUT of the barrel), so it installs `globalThis.OneNote`
 * + `globalThis.Office` itself (mirroring the self-contained installer in
 * `packages/bridge-onenote/src/onenote-bridge.test.ts`) and hands back a `restore()`.
 *
 * Enumerated host calls modelled (the fidelity boundary for OneNote — confirmed against
 * `onenote-bridge.ts`):
 *   READ path (resolveContext / captureDocState / searchDocument / listContext):
 *     - `OneNote.run(cb)`.
 *     - `ctx.application.getActivePageOrNull()` → `page.load('id,title')` → `.isNullObject` / `.id` /
 *       `.title`.
 *     - `page.contents.load('items/type')` → `items[].type` (only `'Outline'` carries paragraphs).
 *     - `content.outline.paragraphs.load('items/type')` → `items[].type` (only `'RichText'`).
 *     - `paragraph.richText.load('text')` → `.text`.
 *   WRITE path (actuate `append-page` — the page-synthesis verb):
 *     - `ctx.application.getActiveSectionOrNull()` → `section.load('id')` → `.isNullObject` / `.id`.
 *     - `section.addPage(title)` → a new `Page`; `page.addOutline(left, top, html)` (the citation-
 *       tagged outline WRITE); `page.load('id')` → `.id`.
 *   CAPABILITY gate:
 *     - `Office.context.requirements.isSetSupported('OneNoteApi', '1.1')` (numeric-version compare,
 *       matching the host's real semantics).
 *
 * Fidelity notes / boundary:
 *   - `addPage` records the new page (title + the outline HTML it receives) into the active section's
 *     `addedPages` log so a test can read the synthesized page back from the host and inspect the
 *     citation-tagged blocks the bridge inserted.
 *   - There is no OneNote object-model event API in this set, so `watch` is not modelled (the bridge
 *     omits it). The bridge never mutates the source page, so the read fixture is immutable here.
 */

/* ─────────────────────────── seed shapes ─────────────────────────── */

/** A page paragraph: `'RichText'` (carries `text`) or any other type (skipped by the bridge). */
export interface OneNoteParagraphSeed {
  type: string;
  text?: string;
}

/** A page content block: `'Outline'` (has paragraphs) or any other type (skipped). */
export interface OneNoteContentSeed {
  type: string;
  paragraphs?: OneNoteParagraphSeed[];
}

/** The active page: id + title + ordered content blocks. */
export interface OneNotePageSeed {
  id: string;
  title: string;
  contents: OneNoteContentSeed[];
}

/** One page appended during a write — the title + the citation-tagged outline HTML it received. */
export interface AddedPageRecord {
  id: string;
  title: string;
  outlineHtml: string | undefined;
  outlineLeft: number | undefined;
  outlineTop: number | undefined;
}

/** The active section: its id + the pages added during the session (the write log). */
export interface OneNoteSectionSeed {
  id: string;
  addedPages: AddedPageRecord[];
}

/** The OneNote notebook seed: an active page (or null) + an active section (or null). */
export interface OneNoteSeed {
  /** The active page, or `null` to simulate "no active page". */
  page: OneNotePageSeed | null;
  /** The active section, or `null` to simulate "no active section" (write degrades). */
  section: OneNoteSectionSeed | null;
}

/* ─────────────────────────── fake object model ─────────────────────────── */

class FakeRichText {
  text: string | undefined = undefined;
  constructor(private readonly seed: OneNoteParagraphSeed) {}
  load(_p?: string): this {
    // The bridge reads `.text` only after a sync; mirror the seed at load time (as Office does).
    this.text = this.seed.text;
    return this;
  }
}

class FakeParagraph {
  readonly richText: FakeRichText;
  constructor(private readonly seed: OneNoteParagraphSeed) {
    this.richText = new FakeRichText(seed);
  }
  get type(): string {
    return this.seed.type;
  }
}

class FakeParagraphCollection {
  items: FakeParagraph[] = [];
  constructor(private readonly seed: OneNoteParagraphSeed[]) {}
  load(_p?: string): this {
    this.items = this.seed.map((p) => new FakeParagraph(p));
    return this;
  }
}

class FakeOutline {
  readonly paragraphs: FakeParagraphCollection;
  constructor(seed: OneNoteParagraphSeed[]) {
    this.paragraphs = new FakeParagraphCollection(seed);
  }
}

class FakePageContent {
  readonly outline: FakeOutline;
  constructor(private readonly seed: OneNoteContentSeed) {
    this.outline = new FakeOutline(seed.paragraphs ?? []);
  }
  get type(): string {
    return this.seed.type;
  }
}

class FakePageContentCollection {
  items: FakePageContent[] = [];
  constructor(private readonly seed: OneNoteContentSeed[]) {}
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
  outlineCalls: Array<{ left: number; top: number; html: string }> = [];
  constructor(private readonly seed: OneNotePageSeed | null) {
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
  /** The page returned by the last `addPage`, so a test may inspect its outline directly. */
  lastAddedPage: FakePage | undefined;
  private counter = 0;
  constructor(private readonly seed: OneNoteSectionSeed | null) {
    this.isNullObject = seed === null;
  }
  load(props?: string): this {
    if (this.seed && props?.includes('id')) this.id = this.seed.id;
    return this;
  }
  addPage(title: string): FakePage {
    if (!this.seed) throw new Error('fake-onenote: addPage on a null section');
    this.counter += 1;
    const id = `${this.seed.id}-page-${this.counter}`;
    const page = new FakePage({ id, title, contents: [] });
    const record: AddedPageRecord = {
      id,
      title,
      outlineHtml: undefined,
      outlineLeft: undefined,
      outlineTop: undefined,
    };
    this.seed.addedPages.push(record);
    // Mirror the outline write back into the record so a test reads it from the seed (the host).
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
  constructor(seed: OneNoteSeed) {
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
}

class FakeContext {
  readonly application: FakeApplication;
  syncs = 0;
  constructor(seed: OneNoteSeed) {
    this.application = new FakeApplication(seed);
  }
  sync(): Promise<void> {
    this.syncs += 1;
    return Promise.resolve();
  }
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** A read-back view of the section's appended pages after a synthesis run. */
export interface OneNoteSnapshot {
  /** Pages appended this session, in order, with the citation-tagged outline HTML they received. */
  addedPages: ReadonlyArray<AddedPageRecord>;
  /** The active page's title (the synthesis source), unchanged by writes. */
  sourceTitle: string | undefined;
}

/** The installed OneNote simulator. */
export interface OneNoteSimulator {
  readonly seed: OneNoteSeed;
  snapshot(): OneNoteSnapshot;
  restore(): void;
}

/** Default requirement support: OneNoteApi 1.1 present (numeric-version compare). */
const DEFAULT_REQUIREMENTS: Record<string, number> = { OneNoteApi: 1.1 };

/**
 * Install an in-memory OneNote host: writes `globalThis.OneNote` + `globalThis.Office` so the REAL
 * {@link "@ge/bridge-onenote"!OneNoteBridge} runs against `seed`. Pass `requirements: { OneNoteApi:
 * 1.0 }` to simulate an older host (the bridge then degrades / reads `[]`).
 */
export function installFakeOneNote(
  seed: OneNoteSeed = defaultOneNoteSeed(),
  requirements: Record<string, number> = DEFAULT_REQUIREMENTS,
): OneNoteSimulator {
  const g = globalThis as Record<string, unknown>;
  const hadOneNote = Object.prototype.hasOwnProperty.call(g, 'OneNote');
  const hadOffice = Object.prototype.hasOwnProperty.call(g, 'Office');
  const prevOneNote = g.OneNote;
  const prevOffice = g.Office;

  g.OneNote = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => cb(new FakeContext(seed)),
  };

  g.Office = {
    context: {
      requirements: {
        isSetSupported: (name: string, version?: string): boolean => {
          const max = requirements[name];
          if (max === undefined) return false;
          const v = version ? Number.parseFloat(version) : 0;
          return v <= max;
        },
      },
    },
  };

  const restore = (): void => {
    if (hadOneNote) g.OneNote = prevOneNote;
    else delete g.OneNote;
    if (hadOffice) g.Office = prevOffice;
    else delete g.Office;
  };

  return {
    seed,
    snapshot: () => ({
      addedPages: (seed.section?.addedPages ?? []).map((p) => ({ ...p })),
      sourceTitle: seed.page?.title,
    }),
    restore,
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** An `'Outline'` content block from rich-text paragraph seeds. */
export function oneNoteOutline(paragraphs: OneNoteParagraphSeed[]): OneNoteContentSeed {
  return { type: 'Outline', paragraphs };
}

/** Build a {@link OneNoteSeed} from an active page + (optionally) an active section. */
export function oneNoteSeed(init: {
  page: OneNotePageSeed | null;
  section?: OneNoteSectionSeed | null;
}): OneNoteSeed {
  return { page: init.page, section: init.section ?? null };
}

/** An empty active section (id + an empty write log) ready to receive synthesized pages. */
export function oneNoteSection(id = 'sec-research'): OneNoteSectionSeed {
  return { id, addedPages: [] };
}

/**
 * A research-unit-ish OneNote notebook: an active "Source review" page (grounding context) and an
 * empty active section the agent can synthesize a new, citation-tagged page into.
 */
export function defaultOneNoteSeed(): OneNoteSeed {
  return oneNoteSeed({
    page: {
      id: 'pg-source-review',
      title: 'Source review',
      contents: [
        oneNoteOutline([
          { type: 'RichText', text: 'Northwind MSA v3 is the current master agreement.' },
          { type: 'RichText', text: 'ISO 27001 certificate is valid through Nov 2026.' },
        ]),
      ],
    },
    section: oneNoteSection(),
  });
}
