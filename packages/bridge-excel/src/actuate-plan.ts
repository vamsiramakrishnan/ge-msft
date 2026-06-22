import type { ActuationRequest, SourceRef } from '@ge/contracts';

/**
 * Pure translation of an actuation into a host plan — testable without Office.js. A
 * `write-cells` is located by an explicit `target.range` (e.g. "Sheet1!A1:B3"); the grid to
 * write is `params.cells`. The bridge parses the address and writes `range.values` at
 * apply-time.
 */
export interface WriteCellsPlan {
  address?: string;
  values: string[][];
}

export function planWriteCells(req: ActuationRequest): WriteCellsPlan {
  const p = req.params;
  return {
    ...(p.target?.range ? { address: p.target.range } : {}),
    values: p.cells ?? [],
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
  /** True iff at least one format facet is set — the bridge degrades a no-op format. */
  hasOps: boolean;
}

export function planFormatCells(req: ActuationRequest): FormatCellsPlan {
  const p = req.params;
  const f = p.format ?? {};
  const hasOps =
    f.bold !== undefined ||
    f.italic !== undefined ||
    f.fill !== undefined ||
    f.numberFormat !== undefined;
  return {
    ...(p.target?.range ? { address: p.target.range } : {}),
    ...(f.bold !== undefined ? { bold: f.bold } : {}),
    ...(f.italic !== undefined ? { italic: f.italic } : {}),
    ...(f.fill !== undefined ? { fill: f.fill } : {}),
    ...(f.numberFormat !== undefined ? { numberFormat: f.numberFormat } : {}),
    hasOps,
  };
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
