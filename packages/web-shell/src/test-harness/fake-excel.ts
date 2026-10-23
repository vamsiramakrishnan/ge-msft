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
 *   - `sheet.tables.add(address, hasHeaders)` (WRITE) → `.load('name')` (host-minted `TableN`,
 *     readable only after sync) / `.getRange()`.
 *   - `sheet.charts.add(type, sourceRange, seriesBy)` (WRITE) → `.title.text=`, `.load('name')`
 *     (minted `Chart N`), `.series.getItemAt(i)` / `.series.add(name)` → `.setValues` /
 *     `.setXAxisValues`, `.delete()`. The source is modelled as ONE implicit series (no seriesBy
 *     splitting into per-column series).
 *   - `range.conditionalFormats` → `.getCount()` (ClientResult, `.value` after sync), `.add(type)`
 *     (WRITE, inserted at the TOP priority like the host) → `.cellValue.rule=` /
 *     `.cellValue.format.fill.color=`, `.topBottom.rule=` / `.topBottom.format.fill.color=`.
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
  /** Review → Protect Sheet is on (every cell locked, the Excel default). */
  protected?: boolean;
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
  /** Set for a table created via `tables.add(address, hasHeaders)`. */
  hasHeaders?: boolean;
}

/** One chart series: the A1 of its values / category (x-axis) ranges, plus its name if given. */
export interface ChartSeriesSeed {
  name?: string;
  /** Sheet-qualified A1 of the series values. */
  values?: string;
  /** Sheet-qualified A1 of the category (x-axis) values. */
  xValues?: string;
}

/** A chart created via `sheet.charts.add(type, sourceRange, seriesBy)`. */
export interface ChartSeed {
  /** Host-minted name, e.g. `"Chart 1"`. */
  name: string;
  sheet: string;
  /** The `Excel.ChartType` string, e.g. `"ColumnClustered"`. */
  type: string;
  /** Sheet-qualified A1 of the `sourceData` range passed to `add`. */
  source: string;
  seriesBy?: string;
  title?: string;
  /** Series in host order; `series[0]` is the implicit series over `source`. */
  series: ChartSeriesSeed[];
}

/** One conditional-format rule added via `range.conditionalFormats.add(type)`. */
export interface ConditionalFormatSeed {
  /** The `Excel.ConditionalFormatType` string, e.g. `"CellValue"`, `"DataBar"`. */
  cfType: string;
  cellValue?: { operator: string; formula1: string; formula2?: string; fill?: string };
  topBottom?: { rank: number; type: string; fill?: string };
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
  /** Charts created via `sheet.charts.add` (in creation order; `chart.delete()` removes one). */
  charts: ChartSeed[];
  /**
   * Conditional-format rules per sheet-qualified range address, in host priority order (index 0 is
   * the top-priority rule — `add()` inserts there).
   */
  conditionalFormats: Map<string, ConditionalFormatSeed[]>;
}

/** The format facets a `format-cells` effect can set on a range. */
export interface RangeFormatSeed {
  bold?: boolean;
  italic?: boolean;
  fill?: string;
  numberFormat?: string;
  fontColor?: string;
  fontSize?: number;
  horizontalAlignment?: string;
  verticalAlignment?: string;
  wrapText?: boolean;
  borders?: Record<string, Record<string, string>>;
  autofit?: boolean;
}

/** Find a sheet by name in the seed (throws on a typo so a mis-seeded test fails loudly). */
function sheetByName(seed: ExcelSeed, name: string): SheetSeed {
  // An address quotes a name with spaces: 'Branch Performance'!A1 (and doubles a quote inside).
  const quoted = /^'(.*)'$/.exec(name);
  const plain = quoted ? quoted[1]!.replace(/''/g, "'") : name;
  const s = seed.sheets.find((x) => x.name === plain);
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
  /** All tables: seeded ones plus any created via `tables.add` (those carry `hasHeaders`). */
  tables: ReadonlyArray<TableSeed>;
  /** Charts created via `charts.add`. */
  charts: ReadonlyArray<ChartSeed>;
  /** Conditional-format rules keyed by range address, top priority first. */
  conditionalFormats: ReadonlyMap<string, ReadonlyArray<ConditionalFormatSeed>>;
}

/** Event sinks the bridge's `watch()` registers; a test fires these to drive the trigger engine. */
export interface ExcelEvents {
  fireSelectionChanged(address: string): void;
  fireChanged(source?: string): void;
  fireCommentAdded(commentId: string, source?: string): void;
}

/* ─────────────────────────── the fake object model ─────────────────────── */

/** Anything a `context.sync()` must resolve: queued loads to flush, queued writes to commit. */
interface Syncable {
  flushLoads(): void;
  commit(): void;
}

/** Registers host proxies created this batch so `sync()` flushes/commits them. */
interface Tracker {
  range(range: FakeRange): FakeRange;
  object<T extends Syncable>(obj: T): T;
}

/** Load/sync gating for non-Range proxies (a prop is readable only after `load()` + `sync()`). */
class LoadGate {
  private readonly loaded = new Set<string>();
  private readonly requested = new Set<string>();
  constructor(private readonly kind: string) {}
  load(props?: string): void {
    for (const raw of (props ?? '').split(',')) {
      const name = raw.trim();
      if (name) this.requested.add(name);
    }
  }
  flush(): void {
    for (const p of this.requested) this.loaded.add(p);
    this.requested.clear();
  }
  require(prop: string): void {
    if (!this.loaded.has(prop))
      throw new Error(
        `fake-excel: ${this.kind} property "${prop}" is not loaded — call load('${prop}') then ` +
          `context.sync() before reading it (Office.js PropertyNotLoaded).`,
      );
  }
}

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
  /** ExcelApi 1.7: the contiguous data around this range — the sheet's used range here. */
  getSurroundingRegion(): FakeRange {
    const region = new FakeRange(
      this.seed,
      this.sheetName,
      stripSheet(usedAddress(sheetByName(this.seed, this.sheetName))),
    );
    // Joins the batch like any other range, so its loads resolve at the next sync.
    return this.tracker ? this.tracker.range(region) : region;
  }
  /** Set by {@link registerRange}; lets child proxies (conditional formats) join the batch. */
  tracker: Tracker | undefined;
  get conditionalFormats(): FakeConditionalFormatCollection {
    return new FakeConditionalFormatCollection(this.seed, this.qualifiedAddress(), this.tracker);
  }
  /** Fake-internal (not a host read): the sheet-qualified address, for recording a write's target. */
  qualifiedAddress(): string {
    return `${this.sheetName}!${this.rangeA1}`;
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
  private readonly edges = new Map<string, Record<string, string>>();
  private autofitted = false;
  /** Queued `range.sort.apply` (ExcelApi 1.2), applied to the seed grid at `sync()`. */
  private pendingSort: { key: number; ascending: boolean; hasHeaders: boolean } | undefined;
  readonly sort = {
    apply: (
      fields: Array<{ key: number; ascending?: boolean }>,
      _matchCase?: boolean,
      hasHeaders?: boolean,
    ): void => {
      const field = fields[0];
      if (field)
        this.pendingSort = {
          key: field.key,
          ascending: field.ascending !== false,
          hasHeaders: hasHeaders === true,
        };
    },
  };
  readonly format = {
    protection: { locked: true, load(_props?: string): void {} },
    font: {
      bold: undefined as boolean | undefined,
      italic: undefined as boolean | undefined,
      color: undefined as string | undefined,
      size: undefined as number | undefined,
    },
    fill: { color: undefined as string | undefined },
    horizontalAlignment: undefined as string | undefined,
    verticalAlignment: undefined as string | undefined,
    wrapText: undefined as boolean | undefined,
    borders: {
      getItem: (edge: string): Record<string, string> => {
        if (!this.edges.has(edge)) this.edges.set(edge, {});
        return this.edges.get(edge)!;
      },
    },
    autofitColumns: (): void => {
      this.autofitted = true;
    },
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
      row.map((v) =>
        typeof v !== 'string' ? v : v.startsWith("'") ? v.slice(1) : (asTypedDate(v) ?? v),
      ),
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
    if (this.pendingSort) {
      const { key, ascending, hasHeaders } = this.pendingSort;
      this.pendingSort = undefined;
      const grid = readGrid(sheet, span.startRow, span.startCol, span.rows, span.cols);
      const head = hasHeaders ? grid.slice(0, 1) : [];
      const num = (v: string) => (v.trim() !== '' ? Number(v) : Number.NaN);
      const body = grid.slice(head.length).sort((a, b) => {
        const [x, y] = [a[key] ?? '', b[key] ?? ''];
        const order =
          !Number.isNaN(num(x)) && !Number.isNaN(num(y))
            ? num(x) - num(y)
            : x.localeCompare(y, undefined, { sensitivity: 'base' });
        return ascending ? order : -order;
      });
      [...head, ...body].forEach((row, r) =>
        row.forEach((value, c) => writeCell(sheet, span.startRow + r, span.startCol + c, value)),
      );
      // The host re-reads values after a sort; a later load of this range sees the new order.
      this._valuesCache = undefined;
    }
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
      ...(this.format.font.color !== undefined ? { fontColor: this.format.font.color } : {}),
      ...(this.format.font.size !== undefined ? { fontSize: this.format.font.size } : {}),
      ...(this.format.horizontalAlignment !== undefined
        ? { horizontalAlignment: this.format.horizontalAlignment }
        : {}),
      ...(this.format.verticalAlignment !== undefined
        ? { verticalAlignment: this.format.verticalAlignment }
        : {}),
      ...(this.format.wrapText !== undefined ? { wrapText: this.format.wrapText } : {}),
      ...(this.edges.size > 0 ? { borders: Object.fromEntries(this.edges) } : {}),
      ...(this.autofitted ? { autofit: true } : {}),
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

/** A table created by `sheet.tables.add` — committed into `seed.tables` at the next `sync()`. */
class FakeAddedTable implements Syncable {
  private readonly gate = new LoadGate('Table');
  private record: TableSeed | undefined;
  constructor(
    private readonly seed: ExcelSeed,
    private readonly tracker: Tracker | undefined,
    private readonly sheetName: string,
    private readonly a1: string,
    private readonly hasHeaders: boolean,
  ) {}
  get name(): string {
    this.gate.require('name');
    return this.record!.name;
  }
  load(props?: string): this {
    this.gate.load(props);
    return this;
  }
  getRange(): FakeRange {
    const range = new FakeRange(this.seed, this.sheetName, this.a1);
    return this.tracker ? this.tracker.range(range) : range;
  }
  flushLoads(): void {
    this.gate.flush();
  }
  commit(): void {
    if (this.record) return;
    const taken = new Set(this.seed.tables.map((t) => t.name));
    let n = 1;
    while (taken.has(`Table${n}`)) n++;
    this.record = {
      name: `Table${n}`,
      range: `${this.sheetName}!${this.a1}`,
      hasHeaders: this.hasHeaders,
    };
    this.seed.tables.push(this.record);
  }
}

class FakeWorksheetTableCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName: string,
    private readonly tracker: Tracker | undefined,
  ) {}
  add(address: string, hasHeaders: boolean): FakeAddedTable {
    // A sheet-qualified address lands the table on THAT sheet (like the host).
    const bang = address.lastIndexOf('!');
    const sheetName = bang >= 0 ? unquote(address.slice(0, bang)) : this.sheetName;
    sheetByName(this.seed, sheetName);
    const a1 = (bang >= 0 ? address.slice(bang + 1) : address).replace(/\$/g, '');
    parseA1(a1); // throw on a malformed address, as the host would
    const table = new FakeAddedTable(this.seed, this.tracker, sheetName, a1, hasHeaders);
    return this.tracker ? this.tracker.object(table) : table;
  }
}

class FakeChartSeries {
  constructor(private readonly target: ChartSeriesSeed) {}
  setValues(range: FakeRange): void {
    this.target.values = range.qualifiedAddress();
  }
  setXAxisValues(range: FakeRange): void {
    this.target.xValues = range.qualifiedAddress();
  }
}

class FakeChartSeriesCollection {
  constructor(private readonly series: ChartSeriesSeed[]) {}
  getItemAt(index: number): FakeChartSeries {
    const target = this.series[index];
    if (!target) throw new Error(`fake-excel: ItemNotFound — no chart series at index ${index}`);
    return new FakeChartSeries(target);
  }
  add(name?: string): FakeChartSeries {
    const target: ChartSeriesSeed = name !== undefined ? { name } : {};
    this.series.push(target);
    return new FakeChartSeries(target);
  }
}

/** A chart created by `sheet.charts.add` — committed into `seed.charts` at the next `sync()`. */
class FakeChart implements Syncable {
  private readonly gate = new LoadGate('Chart');
  private committed = false;
  private deleted = false;
  readonly title = { text: undefined as string | undefined };
  readonly series: FakeChartSeriesCollection;
  constructor(
    private readonly seed: ExcelSeed,
    private readonly record: ChartSeed,
  ) {
    this.series = new FakeChartSeriesCollection(record.series);
  }
  get name(): string {
    this.gate.require('name');
    return this.record.name;
  }
  load(props?: string): this {
    this.gate.load(props);
    return this;
  }
  delete(): void {
    this.deleted = true;
  }
  flushLoads(): void {
    this.gate.flush();
  }
  commit(): void {
    if (this.deleted) {
      const i = this.seed.charts.indexOf(this.record);
      if (i >= 0) this.seed.charts.splice(i, 1);
      return;
    }
    if (this.title.text !== undefined) this.record.title = this.title.text;
    if (this.committed) return;
    this.committed = true;
    let n = 1;
    while (this.seed.charts.some((c) => c.name === `Chart ${n}`)) n++;
    this.record.name = `Chart ${n}`;
    this.seed.charts.push(this.record);
  }
}

class FakeChartCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName: string,
    private readonly tracker: Tracker | undefined,
  ) {}
  add(type: string, sourceData: FakeRange, seriesBy?: string): FakeChart {
    const source = sourceData.qualifiedAddress();
    const record: ChartSeed = {
      name: '',
      sheet: this.sheetName,
      type,
      source,
      ...(seriesBy !== undefined ? { seriesBy } : {}),
      series: [{ values: source }],
    };
    const chart = new FakeChart(this.seed, record);
    return this.tracker ? this.tracker.object(chart) : chart;
  }
}

/** The `ClientResult<T>` shape: `.value` is readable only after the next `sync()`. */
class FakeClientResult<T> implements Syncable {
  private resolved = false;
  constructor(private readonly compute: () => T) {}
  private result: T | undefined;
  get value(): T {
    if (!this.resolved)
      throw new Error('fake-excel: ClientResult.value read before context.sync() resolved it.');
    return this.result as T;
  }
  flushLoads(): void {
    if (this.resolved) return;
    this.result = this.compute();
    this.resolved = true;
  }
  commit(): void {}
}

/** A queued conditional-format rule; its sub-objects mirror the host's write-only rule shapes. */
class FakeConditionalFormat implements Syncable {
  private committed = false;
  readonly cellValue = {
    rule: undefined as { formula1: string; formula2?: string; operator: string } | undefined,
    format: { fill: { color: undefined as string | undefined } },
  };
  readonly topBottom = {
    rule: undefined as { rank: number; type: string } | undefined,
    format: { fill: { color: undefined as string | undefined } },
  };
  constructor(
    private readonly seed: ExcelSeed,
    private readonly address: string,
    private readonly cfType: string,
  ) {}
  flushLoads(): void {}
  commit(): void {
    if (this.committed) return;
    this.committed = true;
    const cv = this.cellValue.rule;
    const tb = this.topBottom.rule;
    const cvFill = this.cellValue.format.fill.color;
    const tbFill = this.topBottom.format.fill.color;
    const rule: ConditionalFormatSeed = {
      cfType: this.cfType,
      ...(cv
        ? {
            cellValue: {
              operator: cv.operator,
              formula1: cv.formula1,
              ...(cv.formula2 !== undefined ? { formula2: cv.formula2 } : {}),
              ...(cvFill !== undefined ? { fill: cvFill } : {}),
            },
          }
        : {}),
      ...(tb
        ? {
            topBottom: {
              rank: tb.rank,
              type: tb.type,
              ...(tbFill !== undefined ? { fill: tbFill } : {}),
            },
          }
        : {}),
    };
    const list = this.seed.conditionalFormats.get(this.address) ?? [];
    // `ConditionalFormatCollection.add` inserts at the first/top priority.
    list.unshift(rule);
    this.seed.conditionalFormats.set(this.address, list);
  }
}

class FakeConditionalFormatCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly address: string,
    private readonly tracker: Tracker | undefined,
  ) {}
  getCount(): FakeClientResult<number> {
    const result = new FakeClientResult(
      () => this.seed.conditionalFormats.get(this.address)?.length ?? 0,
    );
    return this.tracker ? this.tracker.object(result) : result;
  }
  add(type: string): FakeConditionalFormat {
    const cf = new FakeConditionalFormat(this.seed, this.address, type);
    return this.tracker ? this.tracker.object(cf) : cf;
  }
}

class FakeWorksheet {
  /** Set by the context's `wrapSheet`; lets `tables`/`charts` proxies join the batch. */
  tracker: Tracker | undefined;
  get tables(): FakeWorksheetTableCollection {
    return new FakeWorksheetTableCollection(this.seed, this.name, this.tracker);
  }
  get charts(): FakeChartCollection {
    return new FakeChartCollection(this.seed, this.name, this.tracker);
  }
  get id(): string {
    return `sheet:${this.name}`;
  }
  readonly visibility = 'Visible';
  get protection(): { protected: boolean; load(_props?: string): void } {
    return { protected: sheetByName(this.seed, this.name).protected === true, load() {} };
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

/**
 * Real Excel reads text written through `values` like typed input: "Jan 2026" becomes a date
 * (stored as its serial number) unless it starts with `'`. Modelled so the simulator catches what
 * real Excel did on 2026-10-17 (a summary's month labels stopped matching the data's text).
 */
const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
function asTypedDate(text: string): string | undefined {
  const m = /^([A-Za-z]{3})[a-z]*[\s-](\d{2}|\d{4})$/.exec(text.trim());
  const month = m ? MONTHS.indexOf(m[1]!.toLowerCase()) : -1;
  if (!m || month < 0) return undefined;
  const year = m[2]!.length === 2 ? 2000 + Number(m[2]) : Number(m[2]);
  const serial = (Date.UTC(year, month, 1) - Date.UTC(1899, 11, 30)) / 86_400_000;
  return String(serial);
}

class FakeWorksheetCollection {
  readonly onChanged = new EventSink<{ source?: string }>();
  readonly onSelectionChanged = new EventSink<{ address: string }>();
  constructor(private readonly seed: ExcelSeed) {}
  getActiveWorksheet(): FakeWorksheet {
    return new FakeWorksheet(this.seed, this.seed.activeSheet);
  }
  get items(): FakeWorksheet[] {
    // Through `getItem`, which the context wraps to track loads until the next sync.
    return this.seed.sheets.map((s) => this.getItem(s.name));
  }
  load(_props?: string): this {
    return this;
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
  load(_props?: string): this {
    return this;
  }
  get items(): Array<{ authorName: string; content: string }> {
    return this.target.replies.map((content) => ({ authorName: '', content }));
  }
}

class FakeComment {
  constructor(private readonly target: CommentSeed) {}
  get id(): string {
    return this.target.id;
  }
  get content(): string {
    return this.target.content;
  }
  get authorName(): string {
    return '';
  }
  /** The anchor cell (`Comment.getLocation()`, ExcelApi 1.10). */
  getLocation(): { address: string; load(props?: string): void } {
    return { address: this.target.cell, load() {} };
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
  constructor(private readonly seed: ExcelSeed) {}
  load(_props?: string): this {
    this.items = this.seed.comments.map((c) => new FakeComment(c));
    return this;
  }
  add(cellAddress: string, content: string): void {
    // Unique across `Excel.run` batches (a fresh collection per context must not reuse ids).
    const taken = new Set(this.seed.comments.map((c) => c.id));
    let n = 1;
    while (taken.has(`sim-comment-${n}`)) n++;
    this.seed.comments.push({
      id: `sim-comment-${n}`,
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
  /** Every proxy handed out this batch; their queued loads flush and writes commit on sync. */
  private readonly touched: Syncable[] = [];
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
function trackRanges(workbook: FakeWorkbook, touched: Syncable[]): FakeWorkbook {
  const tracker: Tracker = {
    range: (range) => registerRange(range, touched),
    object: (obj) => {
      touched.push(obj);
      return obj;
    },
  };
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
    sheet.tracker = tracker;
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
function registerRange(range: FakeRange, touched: Syncable[]): FakeRange {
  touched.push(range);
  range.tracker = {
    range: (r) => registerRange(r, touched),
    object: (obj) => {
      touched.push(obj);
      return obj;
    },
  };
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
      tables: seed.tables.map((t) => ({ ...t })),
      charts: seed.charts.map((c) => ({ ...c, series: c.series.map((x) => ({ ...x })) })),
      conditionalFormats: new Map(
        [...seed.conditionalFormats].map(([k, v]) => [k, v.map((r) => structuredClone(r))]),
      ),
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
    charts: [],
    conditionalFormats: new Map<string, ConditionalFormatSeed[]>(),
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
