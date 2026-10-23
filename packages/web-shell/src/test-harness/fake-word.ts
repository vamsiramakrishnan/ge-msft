/**
 * In-memory **Word host simulator**. Models the slice of the Office.js object model the real
 * {@link "@ge/bridge-word"!OfficeWordHost} drives (the bridge's only un-faked seam), so the REAL
 * bridge runs unchanged against a seeded document — including the load-bearing `body.search`
 * re-resolution and the search→insert tracked-change batch ordering.
 *
 * Enumerated host calls modelled (the fidelity boundary for Word):
 *   - `Word.run(cb)`; `Word.InsertLocation.replace`; `Word.ChangeTrackingMode.trackAll`.
 *   - `ctx.document.getSelection()` → `.text`.
 *   - `ctx.document.body` → `.text`; `.paragraphs` → items `{ text, styleBuiltIn }`.
 *   - `ctx.document.body.search(q, { matchCase })` → `results.items[]` with `.text` and
 *     `.paragraphs.getFirstOrNullObject()` → `.text` / `.isNullObject` (context-hint read).
 *   - `range.insertText(text, location)` (the tracked-change WRITE) — re-resolved at apply-time
 *     against the live body, recorded into the seed and reflected in subsequent searches.
 *   - `range.insertComment(text)` (comments-as-citations WRITE).
 *   - `ctx.document.changeTrackingMode = trackAll` (recorded for assertions).
 *   - `ctx.document.body.getComments()` → items `{ id }`; `comment.reply(text)`; `comment.resolved=`.
 *   - `ctx.document.customXmlParts.add(xml)` (durable provenance WRITE, via `Office`).
 *
 * Fidelity notes / boundary:
 *   - `body.search` is a case-insensitive substring scan over each paragraph's text; a hit's text is
 *     the matched query (as Office returns the matched range), and its first-paragraph context hint
 *     is the containing paragraph. This is faithful enough for anchor choice + drift degradation.
 *   - `insertText(replace)` replaces the FIRST occurrence of the anchor text in the body, so a later
 *     search for the old anchor degrades to drift (the apply-time re-resolution the bridge relies on).
 */

import { installGlobal, composeRestores } from './globals.js';
import {
  makeFakeOffice,
  makeOfficeSeed,
  type OfficeSeed,
  type OfficeHandlerRegistry,
} from './fake-office.js';

/** A document paragraph: its text + the Word built-in style name (drives heading level). */
export interface WordParagraphSeed {
  text: string;
  styleBuiltIn: string;
  /** Test-only: this paragraph sits in a table cell / ends a section / is in a locked control. */
  inTable?: boolean;
  sectionBreak?: boolean;
  lockedControl?: boolean;
}

/** A document comment (the `getComments()` shape). */
export interface WordCommentSeed {
  id: string;
  text: string;
  replies: string[];
  resolved: boolean;
  /** The commented text (`comment.getRange().text`); empty when not given. */
  anchor?: string;
}

/** The Word document seed: selection text + ordered paragraphs + comments. */
export interface WordSeed {
  selectionText: string;
  /** With nothing selected: the index of the paragraph the cursor sits in (default 0). */
  cursorParagraph?: number;
  paragraphs: WordParagraphSeed[];
  comments: WordCommentSeed[];
  /** Recorded: the last `changeTrackingMode` the bridge set (asserts tracked changes were turned on). */
  changeTrackingMode?: string;
  /** Recorded tracked-change inserts: `{ anchor, text }` per applied replace (+ any font set on it). */
  inserts: WordInsertRecord[];
  /** Recorded comments added via `insertComment`: `{ anchor, text }`. */
  addedComments: Array<{ anchor: string; text: string }>;
  /** Recorded native tables inserted via `insertTable` (cell values per row). */
  tables?: string[][][];
}

/** The whole body text — paragraphs joined by newlines (mirrors `body.text`). */
function bodyText(seed: WordSeed): string {
  return seed.paragraphs.map((p) => p.text).join('\n');
}

/* ─────────────────────────── fake object model ─────────────────────────── */

const INSERT_LOCATION = { replace: 'Replace', start: 'Start', end: 'End' } as const;
const CHANGE_TRACKING_MODE = { trackAll: 'TrackAll', off: 'Off' } as const;

/** One `insertText` call, and the font the bridge then set on the returned range. */
export interface WordInsertRecord {
  anchor: string;
  text: string;
  /** Track Changes was on when the write was applied. */
  tracked?: boolean;
  font?: Record<string, unknown>;
}

interface FakeInsertedRange {
  font: Record<string, unknown>;
}

/** The range `insertText` returns: font writes are recorded on the insert. */
function insertedRange(entry: WordInsertRecord): FakeInsertedRange {
  const font = new Proxy<Record<string, unknown>>(
    {},
    {
      set(_target, key, value): boolean {
        entry.font = { ...entry.font, [String(key)]: value };
        return true;
      },
    },
  );
  return { font };
}

class FakeParagraphProxy {
  isNullObject = false;
  text = '';
  constructor(
    text: string | undefined,
    private readonly seed?: WordSeed,
    private readonly index = -1,
  ) {
    if (text === undefined) this.isNullObject = true;
    else this.text = text;
  }
  load(_props?: string): this {
    return this;
  }
  /** WordApi 1.3: the following paragraph, or a null object after the last one. */
  getNextOrNullObject(): FakeParagraphProxy {
    const next = this.seed?.paragraphs[this.index + 1];
    return new FakeParagraphProxy(next?.text, this.seed, this.index + 1);
  }
}

/** A run of whole paragraphs [from, to] (`Paragraph.getRange('Whole').expandTo(...)`). */
class FakeParagraphBlock {
  private inserted = 0;
  constructor(
    private readonly seed: WordSeed,
    private readonly from: number,
    private readonly to: number,
  ) {}
  expandTo(other: FakeParagraphBlock): FakeParagraphBlock {
    return new FakeParagraphBlock(
      this.seed,
      Math.min(this.from, other.from),
      Math.max(this.to, other.to),
    );
  }
  get paragraphs(): { items: WordParagraphSeed[]; load(_props?: string): void } {
    return { items: this.run(), load: () => undefined };
  }
  /** Tables, sections and locked controls the run touches (seeded per paragraph). */
  get tables(): { items: unknown[]; load(_props?: string): void } {
    return { items: this.run().filter((p) => p.inTable), load: () => undefined };
  }
  get sections(): { items: unknown[]; load(_props?: string): void } {
    const breaks = this.run().filter((p) => p.sectionBreak).length;
    return { items: Array.from({ length: breaks + 1 }), load: () => undefined };
  }
  get contentControls(): {
    items: Array<{ cannotDelete: boolean }>;
    load(_props?: string): void;
  } {
    return {
      items: this.run()
        .filter((p) => p.lockedControl)
        .map(() => ({ cannotDelete: true })),
      load: () => undefined,
    };
  }
  /** Paragraph order (WordApi 1.3 `Range.compareLocationWith`). */
  compareLocationWith(other: FakeParagraphBlock): { value: string } {
    return {
      value: this.from < other.from ? 'Before' : this.from === other.from ? 'Equal' : 'After',
    };
  }
  private run(): WordParagraphSeed[] {
    return this.seed.paragraphs.slice(this.from, this.to + 1);
  }
  /** A native table before the run: one paragraph per row, cells tab-separated. */
  insertTable(_rows: number, _cols: number, _location: string, values: string[][]): void {
    insertParagraphsAfter(
      this.seed,
      this.from - 1,
      values.map((row) => row.join('\t')),
    );
    this.inserted += values.length;
    this.seed.tables = [...(this.seed.tables ?? []), values];
  }
  /** Remove the run's paragraphs (tracked when Track Changes is on: recorded as an insert of ''). */
  delete(): void {
    const at = this.from + this.inserted;
    const removed = this.seed.paragraphs.splice(at, this.to - this.from + 1);
    for (const p of removed)
      this.seed.inserts.push({
        anchor: p.text,
        text: '',
        tracked: this.seed.changeTrackingMode === 'TrackAll',
      });
  }
}

class FakeSearchResult {
  text: string;
  /** Set for a find/replace hit: the exact occurrence a write replaces. */
  position: { para: number; start: number; length: number } | undefined;
  /** For a word from `getTextRanges`: its position and the hit's words. */
  private index = 0;
  private words: string[] = [];
  constructor(
    private readonly seed: WordSeed,
    matchText: string,
  ) {
    this.text = matchText;
  }
  load(_props?: string): this {
    return this;
  }
  /** WordApi 1.3: the hit's words (split on spaces, trimmed). A word's write replaces that span. */
  getTextRanges(
    _delimiters: string[],
    _trim: boolean,
  ): { items: FakeSearchResult[]; load(): void } {
    const words = this.text.split(/\s+/).filter(Boolean);
    const items = words.map((word, i) => {
      const r = new FakeSearchResult(this.seed, word);
      r.index = i;
      r.words = words;
      return r;
    });
    return { items, load: () => undefined };
  }
  expandTo(other: FakeSearchResult): FakeSearchResult {
    return new FakeSearchResult(this.seed, this.words.slice(this.index, other.index + 1).join(' '));
  }
  /** WordApi 1.3: the range cut at its paragraph marks (`\r`), the marks trimmed. */
  split(
    _delimiters: string[],
    _trim: boolean,
    _trimSpacing: boolean,
  ): { items: FakeSearchResult[]; load(): void } {
    const parts = this.text.split('\r');
    while (parts.length > 1 && parts.at(-1) === '') parts.pop();
    return {
      items: parts.map((part) => new FakeSearchResult(this.seed, part)),
      load: () => undefined,
    };
  }
  /** WordApi 1.6 `Range.getTrackedChanges` — the simulated body holds no tracked deletions. */
  getTrackedChanges(): { items: Array<{ type: string }>; load(): void } {
    return { items: [], load: () => undefined };
  }
  /** WordApi 1.3: where this hit sits relative to another (paragraph granularity). */
  compareLocationWith(other: FakeSearchResult): { value: string } {
    const [a, b] = [this.anchorIndex(), other.anchorIndex()];
    return { value: a < b ? 'Before' : a === b ? 'Equal' : 'After' };
  }
  get paragraphs(): {
    getFirst(): { getRange(where: string): FakeParagraphBlock };
    getLast(): { getRange(where: string): FakeParagraphBlock } & {
      getRange(_where: string): { insertOoxml(ooxml: string, _loc: string): void };
    };
    getFirstOrNullObject(): FakeParagraphProxy;
    getLast(): { getRange(_where: string): { insertOoxml(ooxml: string, _loc: string): void } };
  } {
    const index = this.anchorIndex();
    const containing = this.seed.paragraphs[index];
    return {
      getFirst: () => ({ getRange: () => new FakeParagraphBlock(this.seed, index, index) }),
      getFirstOrNullObject: () => new FakeParagraphProxy(containing?.text, this.seed, index),
      // Block content goes after the anchor's paragraph, as its own paragraph(s).
      getLast: () => ({
        getRange: () =>
          Object.assign(new FakeParagraphBlock(this.seed, index, index), {
            insertOoxml: (ooxml: string) =>
              insertParagraphsAfter(this.seed, index, ooxmlText(ooxml)),
          }),
      }),
    };
  }
  /** The paragraph style of the match (Word's `Range.style` / `styleBuiltIn`). */
  get style(): string {
    return this.seed.paragraphs[this.anchorIndex()]?.styleBuiltIn ?? 'Normal';
  }
  set style(name: string) {
    const p = this.seed.paragraphs[this.anchorIndex()];
    if (p) p.styleBuiltIn = name;
  }
  set styleBuiltIn(name: string) {
    this.style = name;
  }
  private anchorIndex(): number {
    return this.seed.paragraphs.findIndex((p) =>
      p.text.toLowerCase().includes(this.text.toLowerCase()),
    );
  }
  /** Inline OOXML after the anchor text. */
  insertOoxml(ooxml: string, _location: string): void {
    const text = ooxmlText(ooxml).join(' ');
    this.seed.inserts.push({
      anchor: this.text,
      text,
      tracked: this.seed.changeTrackingMode === 'TrackAll',
    });
    replaceFirst(this.seed, this.text, `${this.text}${text}`);
  }
  /** A native table after the anchor's paragraph: one paragraph per row, cells tab-separated. */
  insertTable(_rows: number, _cols: number, _location: string, values: string[][]): void {
    insertParagraphsAfter(
      this.seed,
      this.anchorIndex(),
      values.map((row) => row.join('\t')),
    );
    this.seed.tables = [...(this.seed.tables ?? []), values];
  }
  /** The tracked-change WRITE: replace the first occurrence of the anchor in the body. */
  insertText(text: string, _location: string): FakeInsertedRange {
    const entry: WordInsertRecord = {
      anchor: this.text,
      text,
      tracked: this.seed.changeTrackingMode === 'TrackAll',
    };
    this.seed.inserts.push(entry);
    const at = this.position;
    const p = at ? this.seed.paragraphs[at.para] : undefined;
    if (at && p) p.text = p.text.slice(0, at.start) + text + p.text.slice(at.start + at.length);
    else replaceFirst(this.seed, this.text, text);
    return insertedRange(entry);
  }
  /** Comments-as-citations WRITE. */
  insertComment(text: string): void {
    this.seed.addedComments.push({ anchor: this.text, text });
    this.seed.comments.push({
      id: `sim-word-comment-${this.seed.comments.length + 1}`,
      text,
      replies: [],
      resolved: false,
    });
  }
}

/** Incremented by every `ctx.sync()`; a proxy's loaded properties exist only after one. */
let syncGeneration = 0;

class FakeSearchResultCollection {
  private readonly hits: FakeSearchResult[];
  private readonly createdAt = syncGeneration;
  constructor(seed: WordSeed, query: string, matchCase: boolean, matchWholeWord?: boolean) {
    // A find/replace search (it always passes matchWholeWord): one hit per occurrence, at its exact
    // position, honouring whole-word and case like Word. "month" never matches inside "monthly".
    if (matchWholeWord !== undefined) {
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const pattern = new RegExp(
        matchWholeWord ? `(?<![\\p{L}\\p{N}_])${escaped}(?![\\p{L}\\p{N}_])` : escaped,
        matchCase ? 'gu' : 'giu',
      );
      this.hits = seed.paragraphs.flatMap((p, para) =>
        [...p.text.matchAll(pattern)].map((m) => {
          const hit = new FakeSearchResult(seed, m[0]);
          hit.position = { para, start: m.index ?? 0, length: m[0].length };
          return hit;
        }),
      );
      return;
    }
    const hay = (s: string): string => (matchCase ? s : s.toLowerCase());
    const needle = matchCase ? query : query.toLowerCase();
    // One hit per paragraph that contains the query (mirrors Office's per-occurrence ranges,
    // collapsed to paragraph granularity — enough for anchor choice + drift).
    this.hits = seed.paragraphs
      .filter((p) => hay(p.text).includes(needle))
      .map(() => new FakeSearchResult(seed, query));
  }
  /** Like Word, `items` is unavailable until a sync: reading it first throws PropertyNotLoaded. */
  get items(): FakeSearchResult[] {
    if (syncGeneration === this.createdAt)
      throw new Error(
        "fake-word: The property 'items' is not available. Call load and context.sync() first.",
      );
    return this.hits;
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeCommentProxy {
  constructor(private readonly target: WordCommentSeed) {}
  get id(): string {
    return this.target.id;
  }
  get content(): string {
    return this.target.text;
  }
  get authorName(): string {
    return 'Reviewer';
  }
  /** WordApi 1.4: the commented range. */
  getRange(): { text: string; load(_props?: string): void } {
    return { text: this.target.anchor ?? '', load: () => undefined };
  }
  /** WordApi 1.4: the thread's replies. */
  get replies(): {
    items: Array<{ authorName: string; content: string }>;
    load(_p?: string): void;
  } {
    return {
      items: this.target.replies.map((content) => ({ authorName: 'Reviewer', content })),
      load: () => undefined,
    };
  }
  reply(text: string): void {
    this.target.replies.push(text);
  }
  set resolved(value: boolean) {
    this.target.resolved = value;
  }
  get resolved(): boolean {
    return this.target.resolved;
  }
}

class FakeCommentCollection {
  items: FakeCommentProxy[];
  constructor(seed: WordSeed) {
    this.items = seed.comments.map((c) => new FakeCommentProxy(c));
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeParagraphCollection {
  items: WordParagraphSeed[];
  constructor(seed: WordSeed) {
    this.items = seed.paragraphs.map((p) => ({ text: p.text, styleBuiltIn: p.styleBuiltIn }));
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeBody {
  constructor(private readonly seed: WordSeed) {}
  get text(): string {
    return bodyText(this.seed);
  }
  load(_props?: string): this {
    return this;
  }
  get paragraphs(): FakeParagraphCollection {
    return new FakeParagraphCollection(this.seed);
  }
  search(
    query: string,
    opts?: { matchCase?: boolean; matchWholeWord?: boolean },
  ): FakeSearchResultCollection {
    return new FakeSearchResultCollection(
      this.seed,
      query,
      opts?.matchCase ?? false,
      opts?.matchWholeWord,
    );
  }
  getComments(): FakeCommentCollection {
    return new FakeCommentCollection(this.seed);
  }
}

class FakeSelection {
  constructor(private readonly seed: WordSeed) {}
  /** WordApi 1.3: the selection's words; a word's write replaces that span in the body. */
  getTextRanges(delimiters: string[], trim: boolean): { items: FakeSearchResult[]; load(): void } {
    return new FakeSearchResult(this.seed, this.seed.selectionText).getTextRanges(delimiters, trim);
  }
  get text(): string {
    return this.seed.selectionText;
  }
  /** `selection.paragraphs.getFirstOrNullObject()`: the paragraph the selection or cursor is in. */
  get paragraphs(): {
    getFirstOrNullObject(): { text: string; isNullObject: boolean; load(): unknown };
  } {
    const p = this.seed.paragraphs[this.anchorIndex()];
    return {
      getFirstOrNullObject: () => ({
        text: p?.text ?? '',
        isNullObject: !p,
        load() {
          return this;
        },
      }),
    };
  }
  private anchorIndex(): number {
    if (!this.seed.selectionText) return this.seed.cursorParagraph ?? 0;
    return this.seed.paragraphs.findIndex((p) =>
      p.text.toLowerCase().includes(this.seed.selectionText.toLowerCase()),
    );
  }
  load(_props?: string): this {
    return this;
  }
  get style(): string {
    return this.paragraph()?.styleBuiltIn ?? 'Normal';
  }
  set style(name: string) {
    const p = this.paragraph();
    if (p) p.styleBuiltIn = name;
  }
  set styleBuiltIn(name: string) {
    this.style = name;
  }
  private paragraph(): WordParagraphSeed | undefined {
    const needle = this.seed.selectionText.toLowerCase();
    return needle
      ? this.seed.paragraphs.find((p) => p.text.toLowerCase().includes(needle))
      : undefined;
  }
  /** Overwrite the selected text (the only location the bridge uses on a selection is replace). */
  insertText(text: string, _location: string): FakeInsertedRange {
    const entry: WordInsertRecord = {
      anchor: this.seed.selectionText,
      text,
      tracked: this.seed.changeTrackingMode === 'TrackAll',
    };
    this.seed.inserts.push(entry);
    replaceFirst(this.seed, this.seed.selectionText, text);
    this.seed.selectionText = text;
    return insertedRange(entry);
  }
  insertOoxml(ooxml: string, location: string): void {
    this.insertText(ooxmlText(ooxml).join('\n'), location);
  }
  insertTable(_rows: number, _cols: number, _location: string, values: string[][]): void {
    const index = this.anchorIndex();
    insertParagraphsAfter(
      this.seed,
      index,
      values.map((row) => row.join('\t')),
    );
    this.seed.tables = [...(this.seed.tables ?? []), values];
  }
}

class FakeCustomXmlPartCollection {
  constructor(private readonly office: OfficeSeed) {}
  add(xml: string): { id: string } {
    this.office.customXmlParts.push(xml);
    return { id: `xml-${this.office.customXmlParts.length}` };
  }
}

class FakeWordDocument {
  private _mode = '';
  readonly body: FakeBody;
  readonly customXmlParts: FakeCustomXmlPartCollection;
  constructor(
    private readonly seed: WordSeed,
    office: OfficeSeed,
  ) {
    this.body = new FakeBody(seed);
    this.customXmlParts = new FakeCustomXmlPartCollection(office);
  }
  getSelection(): FakeSelection {
    return new FakeSelection(this.seed);
  }
  set changeTrackingMode(mode: string) {
    this._mode = mode;
    this.seed.changeTrackingMode = mode;
  }
  get changeTrackingMode(): string {
    return this._mode;
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeWordContext {
  readonly document: FakeWordDocument;
  constructor(seed: WordSeed, office: OfficeSeed) {
    this.document = new FakeWordDocument(seed, office);
  }
  sync(): Promise<void> {
    syncGeneration += 1;
    return Promise.resolve();
  }
}

/** The `Word` namespace object installed onto `globalThis.Word`. */
interface FakeWordNamespace {
  run<T>(callback: (ctx: FakeWordContext) => Promise<T>): Promise<T>;
  InsertLocation: typeof INSERT_LOCATION;
  ChangeTrackingMode: typeof CHANGE_TRACKING_MODE;
}

/** Replace the first paragraph occurrence of `anchor` with `replacement` in the seed body. */
/** The text runs of an OOXML fragment, one entry per paragraph. */
function ooxmlText(ooxml: string): string[] {
  const paragraphs = ooxml
    .split(/<\/w:p>/i)
    .map((p) => [...p.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gi)].map((m) => m[1]).join(''));
  const texts = paragraphs.filter((t) => t.trim());
  return texts.length ? texts : [ooxml.replace(/<[^>]+>/g, '').trim()];
}

/** Insert body paragraphs after `index` (at the end when the anchor was not found). */
function insertParagraphsAfter(seed: WordSeed, index: number, lines: string[]): void {
  const at = index >= 0 ? index + 1 : seed.paragraphs.length;
  seed.paragraphs.splice(at, 0, ...lines.map((text) => ({ text, styleBuiltIn: 'Normal' })));
}

function replaceFirst(seed: WordSeed, anchor: string, replacement: string): void {
  const needle = anchor.toLowerCase();
  for (const p of seed.paragraphs) {
    const idx = p.text.toLowerCase().indexOf(needle);
    if (idx >= 0) {
      p.text = p.text.slice(0, idx) + replacement + p.text.slice(idx + anchor.length);
      return;
    }
  }
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** A read-back view of the Word document after a run. */
export interface WordSnapshot {
  bodyText: string;
  paragraphs: ReadonlyArray<WordParagraphSeed>;
  comments: ReadonlyArray<WordCommentSeed>;
  inserts: ReadonlyArray<WordInsertRecord>;
  addedComments: ReadonlyArray<{ anchor: string; text: string }>;
  changeTrackingMode?: string;
  tables?: ReadonlyArray<string[][]>;
}

/** The installed Word simulator. */
export interface WordSimulator {
  readonly seed: WordSeed;
  readonly office: OfficeSeed;
  readonly officeHandlers: OfficeHandlerRegistry;
  snapshot(): WordSnapshot;
  restore(): void;
}

/**
 * Install an in-memory Word host: writes `globalThis.Word` + `globalThis.Office` so the REAL
 * {@link "@ge/bridge-word"!WordBridge} runs against `seed`. Defaults to {@link defaultWordSeed}
 * (a contract-ish document) + a modern requirement set.
 */
export function installFakeWord(
  seed: WordSeed = defaultWordSeed(),
  requirements: Record<string, number> = { WordApi: 6 },
): WordSimulator {
  const office = makeOfficeSeed(requirements);
  const { office: officeNs, handlers: officeHandlers } = makeFakeOffice(office);

  const word: FakeWordNamespace = {
    run: async <T>(callback: (ctx: FakeWordContext) => Promise<T>): Promise<T> =>
      callback(new FakeWordContext(seed, office)),
    InsertLocation: INSERT_LOCATION,
    ChangeTrackingMode: CHANGE_TRACKING_MODE,
  };

  const restore = composeRestores([installGlobal('Word', word), installGlobal('Office', officeNs)]);

  return {
    seed,
    office,
    officeHandlers,
    snapshot: () => ({
      bodyText: bodyText(seed),
      paragraphs: seed.paragraphs.map((p) => ({ ...p })),
      ...(seed.tables ? { tables: seed.tables.map((t) => t.map((row) => [...row])) } : {}),
      comments: seed.comments.map((c) => ({ ...c, replies: [...c.replies] })),
      inserts: seed.inserts.map((i) => ({ ...i })),
      addedComments: seed.addedComments.map((c) => ({ ...c })),
      ...(seed.changeTrackingMode !== undefined
        ? { changeTrackingMode: seed.changeTrackingMode }
        : {}),
    }),
    restore,
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** Build a {@link WordSeed} from paragraphs, defaulting empty selection/comments + write logs. */
export function wordSeed(init: {
  selectionText?: string;
  cursorParagraph?: number;
  paragraphs: WordParagraphSeed[];
  comments?: WordCommentSeed[];
}): WordSeed {
  return {
    selectionText: init.selectionText ?? '',
    ...(init.cursorParagraph !== undefined ? { cursorParagraph: init.cursorParagraph } : {}),
    paragraphs: init.paragraphs,
    comments: init.comments ?? [],
    inserts: [],
    addedComments: [],
  };
}

/** A contract-ish Word document fixture: a heading + body paragraphs incl. an SLA claim to redline. */
export function defaultWordSeed(): WordSeed {
  return wordSeed({
    selectionText: 'The SLA is 99.5%.',
    paragraphs: [
      { text: 'Master Services Agreement', styleBuiltIn: 'Heading1' },
      { text: '1. Service Levels', styleBuiltIn: 'Heading2' },
      {
        text: 'The SLA is 99.5%. Downtime beyond this entitles the customer to service credits.',
        styleBuiltIn: 'Normal',
      },
      {
        text: 'Either party may terminate for convenience on 30 days written notice.',
        styleBuiltIn: 'Normal',
      },
    ],
    comments: [
      { id: 'c-existing-1', text: 'Confirm the SLA figure.', replies: [], resolved: false },
    ],
  });
}
