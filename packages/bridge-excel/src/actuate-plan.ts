import {
  gridForRequest,
  type ActuationRequest,
  type CellValue,
  type SourceRef,
} from '@ge/contracts';

/**
 * Pure translation of an actuation into a host plan — testable without Office.js. A
 * `write-cells` is located by an explicit `target.range` (e.g. "Sheet1!A1:B3"); the grid to
 * write uses the shared `cellValues` → legacy `cells` precedence used by verification and
 * recovery. Scalar types stay intact until the bridge encodes literals for Office.js.
 */
export interface WriteCellsPlan {
  address?: string;
  values: CellValue[][];
}

export function planWriteCells(req: ActuationRequest): WriteCellsPlan {
  const p = req.params;
  return {
    ...(p.target?.range ? { address: p.target.range } : {}),
    values: gridForRequest(req),
  };
}

/**
 * Pure plan for a `format-cells` actuation (ADR-0004 Phase 3, the `format` verb). Located by an
 * explicit `target.range`; the host applies ONLY the format facets that are present in
 * `params.format` (each maps to a `range.format.*` write at apply-time), so an empty/absent
 * `format` is a no-op the bridge rejects before touching the host. `hasOps` lets the bridge
 * short-circuit. The fields mirror the {@link ActuationParams.format} schema 1:1; `undefined`
 * means "leave this facet untouched".
 */
export interface FormatCellsPlan {
  address?: string;
  bold?: boolean;
  italic?: boolean;
  /** Background fill color, e.g. "#FFF2CC". */
  fill?: string;
  /** Excel number-format code, e.g. "$#,##0.00". */
  numberFormat?: string;
  fontColor?: string;
  fontSize?: number;
  /** Office.js `HorizontalAlignment` value. */
  horizontalAlignment?: 'Left' | 'Center' | 'Right' | 'Justify';
  /** Office.js `VerticalAlignment` value. */
  verticalAlignment?: 'Top' | 'Center' | 'Bottom';
  wrapText?: boolean;
  /** Applied to all six edges; `None` clears them. */
  borderWeight?: 'Thin' | 'Medium' | 'Thick' | 'None';
  /** Outer-edge weight when it differs from `borderWeight`. */
  outsideWeight?: 'Thin' | 'Medium' | 'Thick' | 'None';
  /** Inner-line weight when it differs from `borderWeight`. */
  insideWeight?: 'Thin' | 'Medium' | 'Thick' | 'None';
  borderColor?: string;
  autofit?: boolean;
  /** True iff at least one format facet is set — the bridge degrades a no-op format. */
  hasOps: boolean;
}

const HORIZONTAL = { left: 'Left', center: 'Center', right: 'Right', justify: 'Justify' } as const;
const VERTICAL = { top: 'Top', center: 'Center', bottom: 'Bottom' } as const;
const WEIGHT = { thin: 'Thin', medium: 'Medium', thick: 'Thick', none: 'None' } as const;

export function planFormatCells(req: ActuationRequest): FormatCellsPlan {
  const p = req.params;
  const f = p.format ?? {};
  const plan: FormatCellsPlan = {
    ...(p.target?.range ? { address: p.target.range } : {}),
    ...(f.bold !== undefined ? { bold: f.bold } : {}),
    ...(f.italic !== undefined ? { italic: f.italic } : {}),
    ...(f.fill !== undefined ? { fill: f.fill } : {}),
    ...(f.numberFormat !== undefined ? { numberFormat: f.numberFormat } : {}),
    ...(f.fontColor !== undefined ? { fontColor: f.fontColor } : {}),
    ...(f.fontSize !== undefined ? { fontSize: f.fontSize } : {}),
    ...(f.align !== undefined ? { horizontalAlignment: HORIZONTAL[f.align] } : {}),
    ...(f.valign !== undefined ? { verticalAlignment: VERTICAL[f.valign] } : {}),
    ...(f.wrap !== undefined ? { wrapText: f.wrap } : {}),
    ...(f.border !== undefined ? { borderWeight: WEIGHT[f.border] } : {}),
    ...(f.borderOutside !== undefined ? { outsideWeight: WEIGHT[f.borderOutside] } : {}),
    ...(f.borderInside !== undefined ? { insideWeight: WEIGHT[f.borderInside] } : {}),
    ...(f.borderColor !== undefined ? { borderColor: f.borderColor } : {}),
    ...(f.autofit ? { autofit: true } : {}),
    hasOps: false,
  };
  plan.hasOps = Object.keys(plan).some((k) => k !== 'address' && k !== 'hasOps');
  return plan;
}

/**
 * Pure plan for an `add-comment` actuation (ADR-0004 `comment` verb on Excel): a new cell-anchored
 * comment whose body is `params.text`, attached to the anchor (first) cell of `target.range`.
 * `text` is collapsed to a single line (a model/host-derived comment is untrusted) so it can't
 * forge structure; `hasTarget`/`hasText` let the bridge reject a malformed request before the host.
 */
export interface AddCommentPlan {
  address?: string;
  text: string;
  hasTarget: boolean;
  hasText: boolean;
}

export function planAddComment(req: ActuationRequest): AddCommentPlan {
  const range = req.params.target?.range;
  const text = oneLineSource(req.params.text ?? '');
  return {
    ...(range ? { address: range } : {}),
    text,
    hasTarget: Boolean(range),
    hasText: text.length > 0,
  };
}

/**
 * The result of routing a written grid into formula vs. value cells (ADR-0003 element 3,
 * "formula-first writes"). Any cell whose string starts with `=` is an Excel formula the user
 * can inspect, so it goes into `formulas` (Excel evaluates it); everything else stays a literal
 * `value`. The two grids are the same shape as the input; the cell NOT taken in a given grid is
 * `null` (Office.js treats a `null` formula/value as "leave this cell untouched"). When the grid
 * has no `=`-prefixed cells, `hasFormulas` is false and the bridge keeps the existing
 * `range.values` write path unchanged.
 */
export interface FormulaGrid {
  formulas: (string | null)[][];
  values: (string | number | boolean | null)[][];
  hasFormulas: boolean;
  /**
   * Formula cells rejected as unsafe to evaluate (see {@link isUnsafeFormula}). A non-empty list
   * means the bridge must degrade the whole write rather than evaluate untrusted active content.
   */
  rejected: string[];
}

/** True iff a written cell is an Excel formula (a string beginning with `=`). */
function isFormula(cell: string): boolean {
  return cell.startsWith('=');
}

/**
 * Active-content / exfiltration functions we must never evaluate when the cell text is
 * model/host-derived (untrusted): web + data fetchers, real-time data, the macro escape hatches,
 * and image/file I/O. Mirrors the set Claude's own Excel add-in gates behind a confirmation. A
 * formula tripping this (or carrying a DDE `|` payload or an external-workbook reference) is
 * rejected and the write degrades, rather than being silently evaluated. See ADR-0003 §untrusted
 * boundary — untrusted content must stay data, never become an executable instruction.
 */
const UNSAFE_FORMULA_FN =
  /\b(WEBSERVICE|STOCKHISTORY|IMPORTDATA|IMPORTXML|IMPORTHTML|IMPORTRANGE|IMPORTFEED|RTD|CALL|EVALUATE|EXEC|FILES|FOPEN|FWRITE|FREAD|IMAGE)\s*\(/i;
const EXTERNAL_WORKBOOK_REF = /\[[^\]]*\.(xls|xlsx|xlsm|xlsb|csv)/i;

/** True iff an `=`-formula must not be evaluated because it is an untrusted active-content vector. */
export function isUnsafeFormula(cell: string): boolean {
  if (UNSAFE_FORMULA_FN.test(cell)) return true;
  if (cell.includes('|')) return true; // DDE-style payload, e.g. =cmd|'/c calc'!A1
  if (EXTERNAL_WORKBOOK_REF.test(cell)) return true;
  return false;
}

/**
 * Route a string grid into parallel formula/value grids. Pure, so it's unit-tested. A *safe*
 * formula cell (`=…`) lands in `formulas` with `null` in `values`; a literal lands in `values`
 * with `null` in `formulas`. An `=`-cell flagged by {@link isUnsafeFormula} is NOT evaluated —
 * it goes to neither grid (both `null`) and is recorded in `rejected` so the bridge degrades the
 * write. `hasFormulas` reports whether any *safe* formula was found — the bridge uses it to choose
 * between the `range.formulas` and `range.values` write paths so non-formula writes are unchanged.
 */
export function splitFormulaGrid(values: string[][]): FormulaGrid {
  let hasFormulas = false;
  const rejected: string[] = [];
  const formulas: (string | null)[][] = [];
  const valueGrid: (string | number | boolean | null)[][] = [];
  for (const row of values) {
    const formulaRow: (string | null)[] = [];
    const valueRow: (string | number | boolean | null)[] = [];
    for (const cell of row) {
      if (isFormula(cell) && isUnsafeFormula(cell)) {
        rejected.push(cell);
        formulaRow.push(null);
        valueRow.push(null);
      } else if (isFormula(cell)) {
        hasFormulas = true;
        formulaRow.push(cell);
        valueRow.push(null);
      } else {
        formulaRow.push(null);
        valueRow.push(cell);
      }
    }
    formulas.push(formulaRow);
    valueGrid.push(valueRow);
  }
  return { formulas, values: valueGrid, hasFormulas, rejected };
}

/* ─────────────────────── ADR-0007 grid-object writes ────────────────────── */

/**
 * Pure plan for a `create-table` actuation (ADR-0007 `table` verb): promote a range to a native
 * Excel Table. Located by `params.table.range`; `hasHeaders` decides whether the first row is the
 * header (Excel synthesizes one when false). `params.table.name` is a *requested* label only — it
 * is NEVER the inverse identity (the bridge records the name Excel actually mints at apply-time;
 * see ADR-0007 §inverse-identity). `hasTable` lets the bridge fail closed before touching the host.
 */
export interface CreateTablePlan {
  address?: string;
  hasHeaders: boolean;
  /** Requested table name (host may rename); informational only, not the inverse identity. */
  requestedName?: string;
  hasTable: boolean;
}

export function planCreateTable(req: ActuationRequest): CreateTablePlan {
  const t = req.params.table;
  const range = t?.range;
  return {
    ...(range ? { address: range } : {}),
    hasHeaders: t?.hasHeaders ?? true,
    ...(t?.name ? { requestedName: t.name } : {}),
    hasTable: Boolean(range),
  };
}

/** The Excel chart-type string this verb supports → the `Excel.ChartType` enum value. */
export type ExcelChartTypeName = 'column' | 'bar' | 'line' | 'pie' | 'scatter' | 'area';
const CHART_TYPE: Record<ExcelChartTypeName, string> = {
  column: 'ColumnClustered',
  bar: 'BarClustered',
  line: 'Line',
  pie: 'Pie',
  scatter: 'XYScatter',
  area: 'Area',
};
const CHART_SERIES_BY: Record<'rows' | 'columns' | 'auto', string> = {
  rows: 'Rows',
  columns: 'Columns',
  auto: 'Auto',
};

/**
 * Pure plan for an `insert-chart` actuation (ADR-0007 `chart` verb): a chart over `sourceRange`.
 * The agent-facing chart type / seriesBy enums are mapped here to their `Excel.ChartType` /
 * `Excel.ChartSeriesBy` string values (host enums are erased strings at runtime), so the host
 * wiring just passes through. `hasChart` lets the bridge fail closed when `sourceRange` is absent.
 */
export interface InsertChartPlan {
  address?: string;
  /** Mapped `Excel.ChartType` string (e.g. "ColumnClustered"). */
  chartType: string;
  /** Mapped `Excel.ChartSeriesBy` string (e.g. "Auto"). */
  seriesBy: string;
  title?: string;
  hasChart: boolean;
}

export function planInsertChart(req: ActuationRequest): InsertChartPlan {
  const c = req.params.chart;
  const sourceRange = c?.sourceRange;
  const chartType = c ? (CHART_TYPE[c.chartType] ?? 'ColumnClustered') : 'ColumnClustered';
  const seriesBy = CHART_SERIES_BY[c?.seriesBy ?? 'auto'];
  return {
    ...(sourceRange ? { address: sourceRange } : {}),
    chartType,
    seriesBy,
    ...(c?.title ? { title: c.title } : {}),
    hasChart: Boolean(sourceRange),
  };
}

/** One area of a chart source: its full address, the address without its first row, its first cell. */
export interface ChartArea {
  address: string;
  body: string;
  topLeft: string;
}

/**
 * Split a chart source into its A1 areas (`Sheet2!C1:C11,Sheet2!G1:G11`). The model asks for this
 * shape to chart non-adjacent columns (categories + values), and `Worksheet.getRange` rejects it, so
 * the bridge builds the chart from the areas instead. All areas must be on one sheet.
 */
export function chartAreas(
  address: string,
): { sheetName?: string; areas: ChartArea[] } | { error: string } {
  const pieces = splitAreaList(address);
  // A lone defined name (`Revenue`, `FY24_Q3`) is one source Excel resolves itself; scanning it for
  // cell addresses read `FY24_Q3` as the areas FY24 and Q3 (security review, 2026-09-30).
  if (pieces.length === 1 && !A1_AREA_RE.test(pieces[0]!) && DEFINED_NAME_RE.test(pieces[0]!))
    return { areas: [{ address: pieces[0]!, body: pieces[0]!, topLeft: pieces[0]! }] };
  const areas: ChartArea[] = [];
  let sheetName: string | undefined;
  for (const piece of pieces) {
    const m = A1_AREA_RE.exec(piece);
    if (!m) return { error: `not an A1 range: ${piece}` };
    const sheet = m[1] === undefined ? undefined : m[1].replace(/^'|'$/g, '').replace(/''/g, "'");
    if (areas.length > 0 && sheet !== sheetName)
      return { error: 'chart areas must all be on one worksheet' };
    sheetName = sheet;
    const firstCol = m[2]!;
    const firstRow = Number(m[3]);
    const lastCol = m[4] ?? firstCol;
    const lastRow = m[5] === undefined ? firstRow : Number(m[5]);
    areas.push({
      address: `${firstCol}${firstRow}:${lastCol}${lastRow}`,
      body: `${firstCol}${Math.min(firstRow + 1, lastRow)}:${lastCol}${lastRow}`,
      topLeft: `${firstCol}${firstRow}`,
    });
  }
  if (areas.length === 0) return { error: `not an A1 range: ${address}` };
  return { ...(sheetName !== undefined ? { sheetName } : {}), areas };
}

/** One whole A1 area, optionally sheet-qualified; anchored so nothing is left over. */
const A1_AREA_RE =
  /^(?:('(?:[^']|'')+'|[^!,\s']+)!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/;
const DEFINED_NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]*$/;

/** Split `A!C1:C9, 'Sales, 2024'!G1:G9` on commas outside quoted sheet names. */
function splitAreaList(address: string): string[] {
  const out: string[] = [];
  let current = '';
  let quoted = false;
  for (const ch of address) {
    if (ch === "'") quoted = !quoted;
    if (ch === ',' && !quoted) {
      out.push(current.trim());
      current = '';
    } else current += ch;
  }
  out.push(current.trim());
  return out.filter(Boolean);
}

const CF_OPERATOR: Record<'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne' | 'between', string> = {
  gt: 'GreaterThan',
  lt: 'LessThan',
  ge: 'GreaterThanOrEqual',
  le: 'LessThanOrEqual',
  eq: 'EqualTo',
  ne: 'NotEqualTo',
  between: 'Between',
};
const CF_TYPE: Record<'cellValue' | 'dataBar' | 'colorScale' | 'top', string> = {
  cellValue: 'CellValue',
  dataBar: 'DataBar',
  colorScale: 'ColorScale',
  top: 'TopBottom',
};

/**
 * Pure plan for a `format-conditional` actuation (ADR-0007 `cf` verb): one conditional-format rule
 * over `range`. The discriminated `rule` is mapped here to its `Excel.ConditionalFormatType` string
 * and (per kind) the host rule shape — a `cellValue` rule carries `formula1`/`formula2?`/`operator`
 * (mapped from the agent operator) plus an optional `fill`; a `top` rule carries `rank`, the
 * `Excel.ConditionalTopBottomCriterionType` string (Top/Bottom × Items), plus an optional `fill`;
 * `dataBar`/`colorScale` are bare. `hasConditional` fails the bridge closed on a missing range.
 */
export type ConditionalCellValuePlan = {
  kind: 'cellValue';
  cfType: string;
  operator: string;
  formula1: string;
  formula2?: string;
  fill?: string;
};
export type ConditionalTopPlan = {
  kind: 'top';
  cfType: string;
  rank: number;
  criterion: string;
  fill?: string;
};
export type ConditionalBarePlan = { kind: 'dataBar' | 'colorScale'; cfType: string };
export type ConditionalRulePlan =
  | ConditionalCellValuePlan
  | ConditionalTopPlan
  | ConditionalBarePlan;

export interface ConditionalPlan {
  address?: string;
  rule?: ConditionalRulePlan;
  hasConditional: boolean;
  /**
   * Security (ADR-0003 §untrusted boundary): a `cellValue` rule whose `formula1`/`formula2` is an
   * untrusted active-content vector (WEBSERVICE/DDE/external-ref/…). Excel evaluates a CF rule's
   * formula, so — exactly like a `=`-cell in {@link splitFormulaGrid} — an unsafe one must NOT be
   * written; the bridge degrades the format instead. Set only for `cellValue`.
   */
  unsafe?: boolean;
}

export function planConditional(req: ActuationRequest): ConditionalPlan {
  const c = req.params.conditional;
  if (!c || !c.range) {
    return { ...(c?.range ? { address: c.range } : {}), hasConditional: false };
  }
  const r = c.rule;
  let rule: ConditionalRulePlan;
  if (r.kind === 'cellValue') {
    rule = {
      kind: 'cellValue',
      cfType: CF_TYPE.cellValue,
      operator: CF_OPERATOR[r.operator],
      formula1: r.value,
      ...(r.value2 !== undefined ? { formula2: r.value2 } : {}),
      ...(r.fill !== undefined ? { fill: r.fill } : {}),
    };
  } else if (r.kind === 'top') {
    rule = {
      kind: 'top',
      cfType: CF_TYPE.top,
      rank: r.rank,
      criterion: r.bottom ? 'BottomItems' : 'TopItems',
      ...(r.fill !== undefined ? { fill: r.fill } : {}),
    };
  } else {
    rule = { kind: r.kind, cfType: CF_TYPE[r.kind] };
  }
  // Screen a cellValue rule's formula(s) the same way `splitFormulaGrid` screens `=`-cells: Excel
  // evaluates a CF formula, so an untrusted active-content payload (WEBSERVICE/DDE/external-ref) must
  // degrade the write rather than be evaluated. `isUnsafeFormula` matches the function/payload by
  // content (a leading `=` is not required), so a bare `WEBSERVICE(...)` threshold is still caught.
  const unsafe =
    rule.kind === 'cellValue' &&
    (isUnsafeFormula(rule.formula1) ||
      (rule.formula2 !== undefined && isUnsafeFormula(rule.formula2)));
  return { address: c.range, rule, hasConditional: true, ...(unsafe ? { unsafe: true } : {}) };
}

/** Collapse control chars + whitespace to one line so a source can't forge extra comment lines. */
function oneLineSource(text: string): string {
  const stripped = Array.from(text, (ch) => {
    const cp = ch.codePointAt(0) ?? 0;
    return cp < 0x20 || cp === 0x7f ? ' ' : ch;
  }).join('');
  return stripped.replace(/\s+/g, ' ').trim();
}

/** Render one source as `Title (uri)`, dropping any non-http(s) uri (no `javascript:`/`data:`). */
export function renderSource(s: SourceRef): string {
  const title = oneLineSource(s.title);
  const uri = s.uri && /^https?:\/\//i.test(s.uri.trim()) ? oneLineSource(s.uri) : undefined;
  return uri ? `${title} (${uri})` : title;
}

/**
 * Render a source list as a plain-text host comment (ADR-0003 element 4,
 * "comments-as-citations"). One `Title (uri)` per line — the uri is dropped when absent or not
 * http(s) — single-lined so a crafted title can't forge extra lines, and capped at `maxChars`
 * (truncated with an ellipsis). An empty source list yields an empty string, which the bridge
 * treats as "no comment to attach".
 */
export function formatSourceComment(sources: SourceRef[], maxChars = 500): string {
  const text = sources.map(renderSource).join('\n');
  if (text.length <= maxChars) return text;
  if (maxChars <= 1) return text.slice(0, maxChars);
  return `${text.slice(0, maxChars - 1)}…`;
}

/** `sort-range` (WS4b): a native in-place sort of one range by one key column. */
export interface SortRangePlan {
  address?: string;
  key: string;
  ascending: boolean;
  hasHeaders: boolean;
}

export function planSortRange(req: ActuationRequest): SortRangePlan {
  const s = req.params.sortRange;
  const address = s?.range?.trim();
  return {
    ...(address ? { address } : {}),
    key: s?.key?.trim() ?? '',
    ascending: s?.order !== 'descending',
    hasHeaders: s?.hasHeaders ?? true,
  };
}

const colToIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0) - 1;

/**
 * The key's column offset within the range: a header name (case-insensitive, exact) wins; else a
 * column letter, relative to the range's first column. Undefined when neither names a column.
 */
export function sortKeyIndex(
  key: string,
  header: readonly unknown[],
  firstColumn: string,
  hasHeaders = true,
): number | undefined {
  // A header name only when row 1 IS a header; a header named "C" wins over column C (review L4).
  if (hasHeaders) {
    const wanted = key.trim().toLowerCase();
    const byName = header.findIndex(
      (cell) =>
        String(cell ?? '')
          .trim()
          .toLowerCase() === wanted,
    );
    if (byName >= 0) return byName;
  }
  if (!/^[A-Za-z]{1,3}$/.test(key.trim())) return undefined;
  const offset = colToIndex(key.trim()) - colToIndex(firstColumn);
  return offset >= 0 && offset < header.length ? offset : undefined;
}

/** A number, or a non-empty string that is one ("98000"); else NaN. */
const asNumber = (v: unknown): number =>
  typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' ? Number(v) : Number.NaN;

const isBlank = (v: unknown): boolean => v === null || v === undefined || v === '';

/** Excel's type order for a sort: numbers, then text, then logicals (blanks handled apart). */
const typeRank = (v: unknown): number =>
  typeof v === 'number' ? 0 : typeof v === 'boolean' ? 2 : 1;

/**
 * True when `a` may come before `b` in an Excel sort with this direction (review H4). Blanks go last
 * in both directions. Two texts that both look like numbers pass in either numeric or text order,
 * since readback may type them either way ("100" < "20" as text; account numbers are text).
 */
function inOrder(a: unknown, b: unknown, ascending: boolean): boolean {
  if (isBlank(b)) return true;
  if (isBlank(a)) return false;
  const dir = ascending ? 1 : -1;
  const [ra, rb] = [typeRank(a), typeRank(b)];
  if (ra !== rb) return (rb - ra) * dir > 0;
  if (ra !== 1) return (Number(b) - Number(a)) * dir >= 0;
  const text = String(b).localeCompare(String(a), undefined, { sensitivity: 'base' }) * dir >= 0;
  const [na, nb] = [asNumber(a), asNumber(b)];
  return text || (!Number.isNaN(na) && !Number.isNaN(nb) && (nb - na) * dir >= 0);
}

/**
 * Built-in functions whose result depends only on their arguments. Anything else (a position or
 * volatile function, a defined name, a LAMBDA or add-in function) makes a formula unsafe to move.
 */
const ROW_SAFE_FUNCTIONS = new Set(
  (
    'ABS AND AVERAGE AVERAGEA CEILING CHAR CONCAT CONCATENATE COUNT COUNTA COUNTBLANK DATE DAY ' +
    'EXACT EXP FIND FLOOR IF IFERROR IFNA IFS INT ISBLANK ISERROR ISNUMBER ISTEXT LEFT LEN LN LOG ' +
    'LOG10 LOWER MAX MAXA MEDIAN MID MIN MINA MOD MONTH N NOT OR POWER PRODUCT PROPER REPT RIGHT ' +
    'ROUND ROUNDDOWN ROUNDUP SEARCH SIGN SQRT SUBSTITUTE SUM SWITCH TEXT TEXTJOIN TRIM TRUNC UPPER ' +
    'VALUE XOR YEAR'
  ).split(' '),
);

const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);

/**
 * True when every formula in the rows a sort moves uses only cells of its own row inside the sorted
 * columns, by relative row, and only built-in functions from {@link ROW_SAFE_FUNCTIONS}. Excel
 * adjusts such references when it moves the row, so each result stays the same (`=E2/D2` →
 * `=E5/D5`). Everything else is refused: another row, a column outside the range, an absolute row
 * (`E$2`), another sheet, a table, spill, trim or whole-column/row reference, a defined name, or any
 * other function. Allow-list: once the known tokens are removed, only operators may remain.
 * `firstRow`/`firstColumn` give the sheet position of `formulas[0][0]`.
 */
export function formulasStayInRow(
  formulas: readonly (readonly unknown[])[],
  firstRow: number,
  firstColumn: string,
  hasHeaders: boolean,
): boolean {
  const left = columnIndex(firstColumn);
  const right = left + (formulas[0]?.length ?? 0) - 1;
  for (let r = hasHeaders ? 1 : 0; r < formulas.length; r++) {
    const row = firstRow + r;
    for (const cell of formulas[r] ?? []) {
      const formula = String(cell ?? '');
      if (!formula.startsWith('=')) continue;
      // String literals are text, not references ("" escapes a quote).
      let code = formula.slice(1).replace(/"(?:[^"]|"")*"/g, ' ');
      if (/[[\]!'#\\@]|:\.|\.:/.test(code)) return false;
      // A whole-row reference (`2:2`, `$3:$9`) spans every column of other rows' data.
      if (/(?<![\w.$])\$?\d+:\$?\d+(?![\w.])/.test(code)) return false;
      let safe = true;
      code = code
        .replace(
          /(?<![\w.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(.])/g,
          (_m, _c: string, col: string, rowLock: string, digits: string) => {
            const at = columnIndex(col);
            if (rowLock || Number(digits) !== row || at < left || at > right) safe = false;
            return ' ';
          },
        )
        .replace(/(?<![\w.])([A-Za-z_][\w.]*)\s*\(/g, (_m, name: string) => {
          if (!ROW_SAFE_FUNCTIONS.has(name.toUpperCase())) safe = false;
          return ' (';
        })
        .replace(/(?<![\w.])(?:TRUE|FALSE)(?![\w.(])/gi, ' ')
        .replace(/(?<![\w.])\d+(?:\.\d+)?(?:E[+-]?\d+)?(?![\w.])/gi, ' ');
      if (!safe || !/^[\s,()+\-*/^&=<>%:]*$/.test(code)) return false;
    }
  }
  return true;
}

/**
 * A sort landed as reviewed when the data rows are the same rows in a new order (nothing changed
 * or lost) and the key column is in the requested order. The header row must not move.
 */
export function sortVerified(
  before: readonly (readonly unknown[])[],
  after: readonly (readonly unknown[])[],
  keyIndex: number,
  ascending: boolean,
  hasHeaders: boolean,
): boolean {
  if (before.length !== after.length) return false;
  const start = hasHeaders ? 1 : 0;
  if (hasHeaders && JSON.stringify(before[0]) !== JSON.stringify(after[0])) return false;
  const rows = (grid: readonly (readonly unknown[])[]) =>
    grid
      .slice(start)
      .map((row) => JSON.stringify(row))
      .sort();
  if (JSON.stringify(rows(before)) !== JSON.stringify(rows(after))) return false;
  for (let r = start + 1; r < after.length; r++)
    if (!inOrder(after[r - 1]![keyIndex], after[r]![keyIndex], ascending)) return false;
  return true;
}
