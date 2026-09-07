/**
 * In-memory **Excel host simulator**. Models the exact slice of the Office.js object model the real
 * {@link "@ge/bridge-excel"!ExcelBridge} drives, so the REAL bridge runs unchanged against seeded
 * workbook data. The proxy graph is materialized in memory, but `Range` reads honor the **load/sync
 * contract**: a property must be named in `load()` and resolved by a `context.sync()` before it can
 * be read (reading early throws, like the real host and `office-addin-mock`). WRITES record back into
 * the seed at `sync()` so a test can assert them via {@link ExcelSimulator.snapshot}.
 *
 * Enumerated host calls modelled (the fidelity boundary for Excel):
 *   - `Excel.run(cb)` — the global entry; runs `cb(ctx)` against a fresh fake `RequestContext`.
 *   - `ctx.workbook.worksheets.getActiveWorksheet()` / `.getItem(name)`.
 *   - `ctx.workbook.getSelectedRange()` → `.address` / `.values`.
 *   - `sheet.getUsedRange()` / `sheet.getUsedRangeOrNullObject()` → `.address` / `.values` /
 *     `.isNullObject` (empty sheet ⇒ null object, ExcelApi 1.4 path).
 *   - `sheet.getRange(a1)` → `.values=` / `.formulas=` (WRITE), `.numberFormat=`,
 *     `.format.font.bold/italic`, `.format.fill.color`, `.getCell(r,c)` → `.address`,
 *     `.rowCount` / `.columnCount` / `.isNullObject` (bounded read metadata).
 *   - `ctx.workbook.tables` → `.load('items/name')` items {name}; `.items[i].getRange()` (table ref
 *     listing and reveal).
 *   - `ctx.workbook.names` → `.load('items/...')` items {name,type,formula}; `.getItemOrNullObject(n)`
 *     → `.getRange()` (named-range read).
 *   - `ctx.workbook.comments` → `.add(cellAddress, content)` (WRITE), `.load('items/id')` items {id},
 *     `.items[i].replies.add(text)` / `.resolved=`, `.onAdded` (event).
 *   - `sheets.onChanged` / `sheets.onSelectionChanged` / `comments.onAdded` — `watch()` events.
 *
 * Out of fidelity scope (stubbed as no-ops, documented so callers know the boundary):
 *   - Cross-sheet formula EVALUATION: a `=`-formula write records the formula string verbatim into
 *     the cell (the bridge's job is to route formulas to `.formulas`, which we record faithfully);
 *     we do NOT recompute dependent cells. The composed-read path (`read | filter | sum`) is computed
 *     by the runtime over the SEEDED values, not by Excel, so this does not affect those tests.
 */

import { parseA1, addressOf, cellRef, indexToCol, type Grid } from './a1.js';
import { installGlobal, composeRestores } from './globals.js';
import {
  makeFakeOffice,
  makeOfficeSeed,
  type OfficeSeed,
  type OfficeHandlerRegistry,
} from './fake-office.js';

/** One named worksheet with a used-range origin and a 2-D string grid of cell values. */
export interface SheetSeed {
  name: string;
  /** Top-left A1 cell of `values` (e.g. `'A1'`). The used range spans from here. */
  origin: string;
  /** Row-major grid of cell values; `''` is an empty cell. */
  values: string[][];
}

/** A workbook-scoped named range (`NamedItem`): a name → its A1 reference (no leading `=`). */
export interface NamedRangeSeed {
  name: string;
  /** Sheet-qualified A1, e.g. `"Sales!A1:D9"`. */
  range: string;
}

/** A workbook-scoped Excel table: a name → its A1 range. */
export interface TableSeed {
  name: string;
  /** Sheet-qualified A1, e.g. `"Sales!A1:D9"`. */
  range: string;
}

/** A cell comment (the `add(cellAddress, content)` shape Excel uses). */
export interface CommentSeed {
  id: string;
  /** Sheet-qualified single-cell address, e.g. `"Sales!F2"`. */
  cell: string;
  content: string;
  replies: string[];
  resolved: boolean;
}

/** The full Excel workbook seed: sheets, the active sheet + selection, names, comments. */
export interface ExcelSeed {
  sheets: SheetSeed[];
  /** Name of the active worksheet (defaults to the first sheet). */
  activeSheet: string;
  /** Sheet-qualified A1 of the current selection, e.g. `"Sales!A2:D2"`. */
  selection: string;
  tables: TableSeed[];
  namedRanges: NamedRangeSeed[];
  comments: CommentSeed[];
  /**
   * Recorded `format-cells` writes, keyed by the sheet-qualified range address the write targeted
   * (e.g. `"Sales!A16:C16"`). Format is write-only on a real `Range` (not part of `.values`), so we
   * record the applied facets here rather than into the value grid — exposed via `snapshot().formats`
   * so a `format-cells` effect is assertable.
   */
  formats: Map<string, RangeFormatSeed>;
}

/** The format facets a `format-cells` effect can set on a range. */
export interface RangeFormatSeed {
  bold?: boolean;
  italic?: boolean;
  fill?: string;
  numberFormat?: string;
}

/** Find a sheet by name in the seed (throws on a typo so a mis-seeded test fails loudly). */
function sheetByName(seed: ExcelSeed, name: string): SheetSeed {
  const s = seed.sheets.find((x) => x.name === name);
  if (!s) throw new Error(`fake-excel: no worksheet named "${name}"`);
  return s;
}

/** The used-range A1 address of a seeded sheet (origin → bottom-right of its grid). */
function usedAddress(sheet: SheetSeed): string {
  const rows = sheet.values.length;
  const cols = Math.max(0, ...sheet.values.map((r) => r.length));
  return addressOf(sheet.name, sheet.origin, rows, cols);
}

/** A read-back view of the workbook after a run, for assertions. */
export interface ExcelSnapshot {
  sheets: ReadonlyArray<{ name: string; values: string[][] }>;
  comments: ReadonlyArray<CommentSeed>;
  /** Recorded `format-cells` writes, keyed by the targeted range address. */
  formats: ReadonlyMap<string, RangeFormatSeed>;
}

/** Event sinks the bridge's `watch()` registers; a test fires these to drive the trigger engine. */
export interface ExcelEvents {
  fireSelectionChanged(address: string): void;
  fireChanged(source?: string): void;
  fireCommentAdded(commentId: string, source?: string): void;
}

/* ─────────────────────────── the fake object model ─────────────────────── */

/** The `Range` read properties subject to load-gating (mirrors office-addin-mock's PropertyNotLoaded). */
const RANGE_READ_PROPS = ['address', 'values', 'rowCount', 'columnCount', 'isNullObject'] as const;

class FakeRange {
  getSpecialCellsOrNullObject() {
    const span = parseA1(this.rangeA1);
    const items: Array<{ address: string }> = [];
    for (let r = 0; r < span.rows; r++)
      for (let c = 0; c < span.cols; c++)
        if (
          (this.pendingFormulas?.[r]?.[c] ||
            (this._pendingValues?.[r]?.[c] === undefined ? this.formulas[r]?.[c] : '')) &&
          String(this.formulas[r]?.[c]).startsWith('=')
        )
          items.push({
            address: `${this.sheetName}!${indexToCol(span.startCol + c)}${span.startRow + r + 1}`,
          });
    return { isNullObject: items.length === 0, load() {}, areas: { items, load() {} } };
  }

  get worksheet(): FakeWorksheet {
    return new FakeWorksheet(this.seed, this.sheetName);
  }
  // Private backings (materialized eagerly from the seed); read access is gated through getters.
  private _isNullObject = false;
  private _address = '';
  private _rowCount = 0;
  private _columnCount = 0;
  /** Lazily materialized read grid — computed on first `.values` read, NOT in the constructor, so a
   * range whose budget check fails (its `.values` is never loaded) never builds the grid. This is
   * what lets the bridge's two-sync read bound run against a huge range without the fake paying for
   * it: `load('rowCount,columnCount')` reads cheap span metadata; only `load('values')` materializes. */
  private _valuesCache: string[][] | undefined;
  /** Queued `.values` write (write side is NOT gated; Office writes don't require a prior load). */
  private _pendingValues: string[][] | undefined;
  // Write-only facets the bridge sets (never read back by the bridge), so left ungated.
  private pendingFormulas?: unknown[][];
  get formulas(): unknown[][] {
    return this.pendingFormulas ?? this.values;
  }
  set formulas(value: unknown[][]) {
    this.pendingFormulas = value;
  }
  numberFormat: unknown[][] = [];
  readonly format = {
    font: { bold: undefined as boolean | undefined, italic: undefined as boolean | undefined },
    fill: { color: undefined as string | undefined },
  };

  /**
   * Office.js proxy fidelity: a property is unreadable until it has been named in `load()` AND a
   * `context.sync()` has resolved it. `requested` holds names from `load()` this batch; `sync()`
   * calls {@link flushLoads} to promote them into `loaded`. Reading an unloaded prop throws, exactly
   * like the real host (and Microsoft's `office-addin-mock`) — so an integration test can't pass
   * against a bridge that reads `.values` without loading it first.
   */
  private readonly loaded = new Set<string>();
  private readonly requested = new Set<string>();

  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName: string,
    private readonly rangeA1: string,
    nullObject = false,
  ) {
    this._isNullObject = nullObject;
    if (!nullObject) this.materialize();
  }

  private materialize(): void {
    const span = parseA1(this.rangeA1);
    this._address = `${this.sheetName}!${this.rangeA1}`;
    this._rowCount = span.rows; // cheap span metadata; the value grid is deferred (see _valuesCache).
    this._columnCount = span.cols;
  }

  /** Build the read grid from the seed on demand (cached). Only reached via the `.values` getter. */
  private computeValues(): string[][] {
    const sheet = sheetByName(this.seed, this.sheetName);
    const span = parseA1(this.rangeA1);
    return readGrid(sheet, span.startRow, span.startCol, span.rows, span.cols);
  }

  private requireLoaded(prop: string): void {
    if (!this.loaded.has(prop))
      throw new Error(
        `fake-excel: property "${prop}" is not loaded — call range.load('${prop}') then ` +
          `context.sync() before reading it (Office.js PropertyNotLoaded).`,
      );
  }

  get isNullObject(): boolean {
    this.requireLoaded('isNullObject');
    return this._isNullObject;
  }
  get address(): string {
    this.requireLoaded('address');
    return this._address;
  }
  get rowCount(): number {
    this.requireLoaded('rowCount');
    return this._rowCount;
  }
  get columnCount(): number {
    this.requireLoaded('columnCount');
    return this._columnCount;
  }
  get values(): string[][] {
    this.requireLoaded('values');
    return this._pendingValues ?? (this._valuesCache ??= this.computeValues());
  }
  set values(grid: unknown[][]) {
    this._pendingValues = grid.map((row) =>
      row.map((v) => (typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v)),
    ) as string[][];
  }

  /** Queue a property (or comma-list, or all when omitted) for resolution at the next `sync()`. */
  load(props?: string): this {
    const names =
      props && props.trim() ? props.split(',') : (RANGE_READ_PROPS as readonly string[]);
    for (const raw of names) {
      const name = raw.trim().split('/')[0]?.trim();
      if (name) this.requested.add(name);
    }
    return this;
  }

  /** Promote this batch's requested loads into readable props (called by `context.sync()`). */
  flushLoads(): void {
    for (const p of this.requested) this.loaded.add(p);
    this.requested.clear();
  }

  getCell(rowOffset: number, colOffset: number): FakeRange {
    const span = parseA1(this.rangeA1);
    const cellA1 = cellRef(span.startCol + colOffset, span.startRow + rowOffset);
    return new FakeRange(this.seed, this.sheetName, cellA1);
  }

  /**
   * Commit any queued `values`/`formulas`/`format` writes back into the seed (called at `sync()`).
   * Uses the private backings directly — write-back is internal bookkeeping, not a gated host read.
   */
  commit(): void {
    if (this._isNullObject) return;
    this.commitFormat();
    const sheet = sheetByName(this.seed, this.sheetName);
    const span = parseA1(this.rangeA1);
    // A formula write takes precedence per cell (Excel routes `=` cells to `.formulas`); literal
    // cells fall through to `.values`. We record the resolved string into the seed grid.
    for (let r = 0; r < span.rows; r++) {
      for (let c = 0; c < span.cols; c++) {
        const formula = this.pendingFormulas?.[r]?.[c];
        const value = this._pendingValues?.[r]?.[c];
        const written =
          formula !== undefined && formula !== null && formula !== ''
            ? String(formula)
            : value !== undefined
              ? String(value)
              : undefined;
        if (written !== undefined) writeCell(sheet, span.startRow + r, span.startCol + c, written);
      }
    }
  }

  /** Record any queued format facets against this range's address (format is write-only on Range). */
  private commitFormat(): void {
    const numberFormat = this.numberFormat[0]?.[0];
    const facets: RangeFormatSeed = {
      ...(this.format.font.bold !== undefined ? { bold: this.format.font.bold } : {}),
      ...(this.format.font.italic !== undefined ? { italic: this.format.font.italic } : {}),
      ...(this.format.fill.color !== undefined ? { fill: this.format.fill.color } : {}),
      ...(numberFormat !== undefined ? { numberFormat: String(numberFormat) } : {}),
    };
    if (Object.keys(facets).length === 0) return;
    const prev = this.seed.formats.get(this._address) ?? {};
    this.seed.formats.set(this._address, { ...prev, ...facets });
  }
}

class FakeNamedItem {
  constructor(
    private readonly seed: ExcelSeed,
    readonly name: string,
    readonly type: string,
    readonly formula: string,
    readonly isNullObject = false,
  ) {}
  load(_props?: string): this {
    return this;
  }
  getRange(): FakeRange {
    if (this.isNullObject) return new FakeRange(this.seed, this.seed.activeSheet, 'A1', true);
    const ref = this.formula.replace(/^=/, '');
    const bang = ref.lastIndexOf('!');
    const sheetName = bang >= 0 ? unquote(ref.slice(0, bang)) : this.seed.activeSheet;
    const a1 = bang >= 0 ? ref.slice(bang + 1) : ref;
    return new FakeRange(this.seed, sheetName, a1.replace(/\$/g, ''));
  }
}

class FakeTable {
  constructor(
    private readonly seed: ExcelSeed,
    readonly name: string,
    readonly range: string,
  ) {}
  getRange(): FakeRange {
    const bang = this.range.lastIndexOf('!');
    const sheetName = bang >= 0 ? unquote(this.range.slice(0, bang)) : this.seed.activeSheet;
    const a1 = bang >= 0 ? this.range.slice(bang + 1) : this.range;
    return new FakeRange(this.seed, sheetName, a1.replace(/\$/g, ''));
  }
}

class FakeTableCollection {
  items: FakeTable[] = [];
  constructor(private readonly seed: ExcelSeed) {}
  load(_props?: string): this {
    this.items = this.seed.tables.map((t) => new FakeTable(this.seed, t.name, t.range));
    return this;
  }
}

class FakeWorksheet {
  get id(): string {
    return `sheet:${this.name}`;
  }
  constructor(
    private readonly seed: ExcelSeed,
    readonly name: string,
  ) {}
  load(_props?: string): this {
    return this;
  }
  getUsedRange(): FakeRange {
    const sheet = sheetByName(this.seed, this.name);
    return new FakeRange(this.seed, this.name, stripSheet(usedAddress(sheet)));
  }
  getUsedRangeOrNullObject(): FakeRange {
    const sheet = sheetByName(this.seed, this.name);
    const empty = sheet.values.every((row) => row.every((c) => String(c ?? '').trim() === ''));
    if (empty) return new FakeRange(this.seed, this.name, 'A1', true);
    return new FakeRange(this.seed, this.name, stripSheet(usedAddress(sheet)));
  }
  getRange(a1: string): FakeRange {
    return new FakeRange(this.seed, this.name, a1.replace(/\$/g, ''));
  }
}

/** A registered Office.js event handler with a `.remove()` (the `EventHandlerResult` shape). */
interface EventHandlerResult {
  remove(): void;
}

class EventSink<A> {
  readonly handlers: Array<(args: A) => unknown> = [];
  add(handler: (args: A) => unknown): EventHandlerResult {
    this.handlers.push(handler);
    return {
      remove: () => {
        const i = this.handlers.indexOf(handler);
        if (i >= 0) this.handlers.splice(i, 1);
      },
    };
  }
  fire(args: A): void {
    for (const h of [...this.handlers]) void h(args);
  }
}

class FakeWorksheetCollection {
  readonly onChanged = new EventSink<{ source?: string }>();
  readonly onSelectionChanged = new EventSink<{ address: string }>();
  constructor(private readonly seed: ExcelSeed) {}
  getActiveWorksheet(): FakeWorksheet {
    return new FakeWorksheet(this.seed, this.seed.activeSheet);
  }
  getItem(name: string): FakeWorksheet {
    sheetByName(this.seed, name); // throw on a typo
    return new FakeWorksheet(this.seed, name);
  }
}

class FakeNamedItemCollection {
  items: FakeNamedItem[] = [];
  constructor(private readonly seed: ExcelSeed) {}
  load(_props?: string): this {
    this.items = this.seed.namedRanges.map(
      (n) => new FakeNamedItem(this.seed, n.name, 'Range', `=${n.range}`),
    );
    return this;
  }
  getItemOrNullObject(name: string): FakeNamedItem {
    const found = this.seed.namedRanges.find((n) => n.name === name);
    if (!found) return new FakeNamedItem(this.seed, name, 'Range', '', true);
    return new FakeNamedItem(this.seed, found.name, 'Range', `=${found.range}`);
  }
}

class FakeCommentReplies {
  constructor(private readonly target: CommentSeed) {}
  add(text: string): void {
    this.target.replies.push(text);
  }
}

class FakeComment {
  constructor(private readonly target: CommentSeed) {}
  get id(): string {
    return this.target.id;
  }
  get replies(): FakeCommentReplies {
    return new FakeCommentReplies(this.target);
  }
  set resolved(value: boolean) {
    this.target.resolved = value;
  }
  get resolved(): boolean {
    return this.target.resolved;
  }
}

class FakeCommentCollection {
  items: FakeComment[] = [];
  readonly onAdded = new EventSink<{
    source?: string;
    commentDetails: Array<{ commentId: string }>;
  }>();
  private nextId = 1;
  constructor(private readonly seed: ExcelSeed) {}
  load(_props?: string): this {
    this.items = this.seed.comments.map((c) => new FakeComment(c));
    return this;
  }
  add(cellAddress: string, content: string): void {
    this.seed.comments.push({
      id: `sim-comment-${this.nextId++}`,
      cell: cellAddress,
      content,
      replies: [],
      resolved: false,
    });
  }
}

class FakeWorkbook {
  readonly worksheets: FakeWorksheetCollection;
  readonly tables: FakeTableCollection;
  readonly names: FakeNamedItemCollection;
  readonly comments: FakeCommentCollection;
  constructor(private readonly seed: ExcelSeed) {
    this.worksheets = new FakeWorksheetCollection(seed);
    this.tables = new FakeTableCollection(seed);
    this.names = new FakeNamedItemCollection(seed);
    this.comments = new FakeCommentCollection(seed);
  }
  getSelectedRange(): FakeRange {
    const bang = this.seed.selection.lastIndexOf('!');
    const sheetName = bang >= 0 ? this.seed.selection.slice(0, bang) : this.seed.activeSheet;
    const a1 = bang >= 0 ? this.seed.selection.slice(bang + 1) : this.seed.selection;
    return new FakeRange(this.seed, sheetName, a1.replace(/\$/g, ''));
  }
}

/** The fake `Excel.RequestContext` — tracks ranges so queued writes commit on `sync()`. */
class FakeRequestContext {
  readonly workbook: FakeWorkbook;
  /** Every FakeRange handed out this batch; their queued writes commit on sync. */
  private readonly touched: FakeRange[] = [];
  constructor(seed: ExcelSeed) {
    this.workbook = trackRanges(new FakeWorkbook(seed), this.touched);
  }
  sync(): Promise<void> {
    // Resolve queued loads first (so a property loaded this batch reads back), then commit writes.
    for (const r of this.touched) r.flushLoads();
    for (const r of this.touched) r.commit();
    return Promise.resolve();
  }
}

/**
 * Wrap every `FakeRange` the bridge obtains so its queued `values`/`formulas` writes are committed
 * when the context syncs. Ranges are produced lazily by getters/methods, so we proxy the workbook
 * graph's range-returning calls to register each range into `touched`.
 */
function trackRanges(workbook: FakeWorkbook, touched: FakeRange[]): FakeWorkbook {
  const register = <T>(value: T): T => {
    if (value instanceof FakeRange) touched.push(value);
    return value;
  };
  // Patch the range-producing seams: getSelectedRange, worksheet.getRange/getUsedRange*, getCell.
  const origSelected = workbook.getSelectedRange.bind(workbook);
  workbook.getSelectedRange = () => register(origSelected());

  const origGetItem = workbook.worksheets.getItem.bind(workbook.worksheets);
  const origActive = workbook.worksheets.getActiveWorksheet.bind(workbook.worksheets);
  const wrapSheet = (sheet: FakeWorksheet): FakeWorksheet => {
    const gr = sheet.getRange.bind(sheet);
    sheet.getRange = (a1: string) => registerRange(gr(a1), touched);
    const gu = sheet.getUsedRange.bind(sheet);
    sheet.getUsedRange = () => registerRange(gu(), touched);
    const gun = sheet.getUsedRangeOrNullObject.bind(sheet);
    sheet.getUsedRangeOrNullObject = () => registerRange(gun(), touched);
    return sheet;
  };
  workbook.worksheets.getItem = (name: string) => wrapSheet(origGetItem(name));
  workbook.worksheets.getActiveWorksheet = () => wrapSheet(origActive());

  // The named-range read path: `names.getItemOrNullObject(name).getRange()` returns a range that
  // must also commit/flush on sync, or a `read <NamedRange>` would read an unloaded property.
  const origTablesLoad = workbook.tables.load.bind(workbook.tables);
  workbook.tables.load = (props?: string): FakeTableCollection => {
    const out = origTablesLoad(props);
    for (const table of workbook.tables.items) {
      const gr = table.getRange.bind(table);
      table.getRange = () => registerRange(gr(), touched);
    }
    return out;
  };

  // The named-range read path: `names.getItemOrNullObject(name).getRange()` returns a range that
  // must also commit/flush on sync, or a `read <NamedRange>` would read an unloaded property.
  const origNamed = workbook.names.getItemOrNullObject.bind(workbook.names);
  workbook.names.getItemOrNullObject = (name: string): FakeNamedItem => {
    const item = origNamed(name);
    const gr = item.getRange.bind(item);
    (item as { getRange: () => FakeRange }).getRange = () => registerRange(gr(), touched);
    return item;
  };
  return workbook;
}

/** Register a range (and wrap its `getCell` so a cell-anchor read/write also commits). */
function registerRange(range: FakeRange, touched: FakeRange[]): FakeRange {
  touched.push(range);
  const gc = range.getCell.bind(range);
  range.getCell = (r: number, c: number) => {
    const cell = gc(r, c);
    touched.push(cell);
    return cell;
  };
  return range;
}

/** The `Excel` namespace object installed onto `globalThis.Excel`. */
interface FakeExcelNamespace {
  run<T>(callback: (ctx: FakeRequestContext) => Promise<T>): Promise<T>;
}

/* ─────────────────────────── grid read/write helpers ───────────────────── */

/** Read a sub-grid of a seeded sheet at a zero-based offset, padding short rows with `''`. */
function readGrid(
  sheet: SheetSeed,
  startRow: number,
  startCol: number,
  rows: number,
  cols: number,
): Grid {
  const origin = parseA1(stripSheet(usedAddressOrigin(sheet)));
  const out: Grid = [];
  for (let r = 0; r < rows; r++) {
    const row: string[] = [];
    for (let c = 0; c < cols; c++) {
      const sr = startRow - origin.startRow + r;
      const sc = startCol - origin.startCol + c;
      row.push(String(sheet.values[sr]?.[sc] ?? ''));
    }
    out.push(row);
  }
  return out;
}

/** Write one cell of a seeded sheet at a zero-based absolute row/col (growing the grid as needed). */
function writeCell(sheet: SheetSeed, absRow: number, absCol: number, value: string): void {
  const origin = parseA1(stripSheet(usedAddressOrigin(sheet)));
  const r = absRow - origin.startRow;
  const c = absCol - origin.startCol;
  if (r < 0 || c < 0) return; // a write above/left of the origin is outside the modelled grid.
  while (sheet.values.length <= r) sheet.values.push([]);
  const row = sheet.values[r] as string[];
  while (row.length <= c) row.push('');
  row[c] = value;
}

/** The sheet's origin cell as a sheet-qualified address (for `parseA1`). */
function usedAddressOrigin(sheet: SheetSeed): string {
  return `${sheet.name}!${sheet.origin}`;
}

/** Drop a `Sheet!` prefix, leaving the bare A1 reference. */
function stripSheet(address: string): string {
  const bang = address.lastIndexOf('!');
  return bang >= 0 ? address.slice(bang + 1) : address;
}

/** Unwrap Excel's `'Sheet Name'` quoting. */
function unquote(name: string): string {
  return name.startsWith("'") && name.endsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** The installed Excel simulator: the seed plus snapshot/events/restore controls. */
export interface ExcelSimulator {
  /** The live seed (mutated by writes); prefer {@link snapshot} for assertions. */
  readonly seed: ExcelSeed;
  /** The Office-level seed (requirements / settings / customXmlParts). */
  readonly office: OfficeSeed;
  /** A read-back view of the workbook after a run. */
  snapshot(): ExcelSnapshot;
  /** Fire the bridge-registered host events to drive the trigger engine. */
  readonly events: ExcelEvents;
  /** Fire raw Office-level events (selection/view), e.g. the Word/PowerPoint Office bus. */
  readonly officeHandlers: OfficeHandlerRegistry;
  /** Remove the installed globals (`Excel`, `Office`). Idempotent. */
  restore(): void;
}

/**
 * Install an in-memory Excel host: writes `globalThis.Excel` + `globalThis.Office` so the REAL
 * {@link "@ge/bridge-excel"!ExcelBridge} runs against `seed`. Defaults to the rich
 * {@link defaultExcelSeed} FSI workbook + a modern requirement set; override either to customize.
 */
export function installFakeExcel(
  seed: ExcelSeed = defaultExcelSeed(),
  requirements: Record<string, number> = { ExcelApi: 13 },
): ExcelSimulator {
  const office = makeOfficeSeed(requirements);
  const { office: officeNs, handlers: officeHandlers } = makeFakeOffice(office);

  // Hold the most-recent context's collections so `events.*` fires the same sinks the bridge wired.
  let lastWorkbook: FakeWorkbook | undefined;
  const excel: FakeExcelNamespace = {
    run: async <T>(callback: (ctx: FakeRequestContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeRequestContext(seed);
      lastWorkbook = ctx.workbook;
      return callback(ctx);
    },
  };

  const restore = composeRestores([
    installGlobal('Excel', excel),
    installGlobal('Office', officeNs),
  ]);

  const events: ExcelEvents = {
    fireSelectionChanged(address) {
      lastWorkbook?.worksheets.onSelectionChanged.fire({ address });
    },
    fireChanged(source) {
      lastWorkbook?.worksheets.onChanged.fire(source !== undefined ? { source } : {});
    },
    fireCommentAdded(commentId, source) {
      lastWorkbook?.comments.onAdded.fire({
        ...(source !== undefined ? { source } : {}),
        commentDetails: [{ commentId }],
      });
    },
  };

  return {
    seed,
    office,
    events,
    officeHandlers,
    snapshot: () => ({
      sheets: seed.sheets.map((s) => ({ name: s.name, values: s.values.map((r) => [...r]) })),
      comments: seed.comments.map((c) => ({ ...c, replies: [...c.replies] })),
      formats: new Map([...seed.formats].map(([k, v]) => [k, { ...v }])),
    }),
    restore,
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** Build an {@link ExcelSeed} from sheets, defaulting active sheet + selection sensibly. */
export function excelSeed(init: {
  sheets: SheetSeed[];
  activeSheet?: string;
  selection?: string;
  tables?: TableSeed[];
  namedRanges?: NamedRangeSeed[];
  comments?: CommentSeed[];
}): ExcelSeed {
  const first = init.sheets[0];
  if (!first) throw new Error('excelSeed: at least one sheet is required');
  const active = init.activeSheet ?? first.name;
  return {
    sheets: init.sheets,
    activeSheet: active,
    selection: init.selection ?? `${active}!${first.origin}`,
    tables: init.tables ?? [],
    namedRanges: init.namedRanges ?? [],
    comments: init.comments ?? [],
    formats: new Map<string, RangeFormatSeed>(),
  };
}

/**
 * A realistic FSI (financial-services) workbook fixture: a `Sales` sheet with a header row + six
 * regional revenue/cost rows, a `Summary` sheet, a named range, and a starting selection over a
 * data row. Mirrors the kind of grid the mockups assume so an integration test reads believable data.
 */
export function defaultExcelSeed(): ExcelSeed {
  return excelSeed({
    sheets: [
      {
        name: 'Sales',
        origin: 'A1',
        values: [
          ['region', 'rep', 'revenue', 'cost'],
          ['East', 'Alice', '300', '120'],
          ['East', 'Bob', '250', '100'],
          ['West', 'Carol', '180', '90'],
          ['West', 'Dan', '220', '110'],
          ['North', 'Erin', '140', '70'],
          ['South', 'Frank', '90', '40'],
        ],
      },
      {
        name: 'Summary',
        origin: 'A1',
        values: [
          ['metric', 'value'],
          ['total revenue', ''],
        ],
      },
    ],
    activeSheet: 'Sales',
    selection: 'Sales!A2:D2',
    tables: [{ name: 'SalesTable', range: 'Sales!A1:D7' }],
    namedRanges: [{ name: 'SalesTable', range: 'Sales!A1:D7' }],
    comments: [],
  });
}
