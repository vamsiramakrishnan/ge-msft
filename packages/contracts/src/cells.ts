import type { ActuationRequest } from './capability.js';
import type { CellSnapshot, CellValue } from './analysis.js';

export function gridForRequest(request: ActuationRequest): CellValue[][] {
  return request.params.cellValues ?? request.params.cells ?? [];
}
export function formulasForRequest(request: ActuationRequest): string[][] {
  const values = gridForRequest(request);
  return values.map((row, r) =>
    row.map((v, c) =>
      request.params.cellValues
        ? (request.params.cellFormulas?.[r]?.[c] ?? '')
        : typeof v === 'string' && v.startsWith('=')
          ? v
          : '',
    ),
  );
}
/** Compare formulas as expressions and values without locale-dependent display formatting. */
export function cellsMatchRequest(snapshot: CellSnapshot, request: ActuationRequest): boolean {
  const values = gridForRequest(request);
  const formulas = formulasForRequest(request);
  if (snapshot.values.length !== values.length || snapshot.values[0]?.length !== values[0]?.length)
    return false;
  return values.every((row, r) =>
    row.every((v, c) => {
      const wantedFormula = formulas[r]?.[c] ?? '';
      const actualFormula = snapshot.formulas?.[r]?.[c] ?? '';
      if (wantedFormula.startsWith('=')) return actualFormula === wantedFormula;
      return (
        !actualFormula.startsWith('=') && String(snapshot.values[r]?.[c] ?? '') === String(v ?? '')
      );
    }),
  );
}
export function hasFormulaErrors(snapshot: CellSnapshot): boolean {
  return snapshot.values.some((row, r) =>
    row.some(
      (v, c) =>
        snapshot.formulas?.[r]?.[c]?.startsWith('=') &&
        /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|N\/A|NUM!|SPILL!|CALC!|NULL!)/.test(String(v)),
    ),
  );
}
export function rangeForGrid(address: string, rows: number, columns: number): string {
  if (
    !Number.isInteger(rows) ||
    !Number.isInteger(columns) ||
    rows < 1 ||
    columns < 1 ||
    rows * columns > 100_000
  )
    throw new Error('Write dimensions exceed the cell budget.');
  const bang = address.lastIndexOf('!');
  const prefix = bang < 0 ? '' : address.slice(0, bang + 1);
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?::\$?([A-Za-z]{1,3})\$?([1-9]\d*))?$/.exec(
    address.slice(bang + 1),
  );
  if (!match) throw new Error('Choose an explicit A1 destination.');
  const index = (s: string): number =>
    [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const firstColumn = index(match[1]!);
  const firstRow = Number(match[2]);
  const lastColumn = firstColumn + columns - 1;
  const lastRow = firstRow + rows - 1;
  if (lastColumn > 16384 || lastRow > 1048576)
    throw new Error('Write exceeds worksheet dimensions.');
  if (match[3] && (index(match[3]) !== lastColumn || Number(match[4]) !== lastRow))
    throw new Error('The destination range dimensions do not match the result.');
  const label = (n: number): string => {
    let s = '';
    while (n) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  };
  const origin = `${label(firstColumn)}${firstRow}`;
  return `${prefix}${origin}${rows === 1 && columns === 1 ? '' : `:${label(lastColumn)}${lastRow}`}`;
}

const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const columnLabel = (index: number): string => {
  let s = '';
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

const MAX_COLUMN = 16384;
const MAX_ROW = 1048576;
// A reference followed by a word character, `.`, `(` (a function), `!` (a sheet name) or `[`
// (a table name) is not a reference.
const CELL_REF = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w.(![])/y;
const COLUMN_SPAN = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w.(![])/y;
const ROW_SPAN = /(\$?)(\d+):(\$?)(\d+)(?![\w.(![])/y;
const WORD = /[\w.\\]+/y;

function matchAt(pattern: RegExp, text: string, at: number): RegExpExecArray | null {
  pattern.lastIndex = at;
  return pattern.exec(text);
}

/** The index after a `"…"` string or `'…'` sheet name that starts at `start` (doubled quote = escape). */
function quotedEnd(text: string, start: number): number {
  const quote = text[start];
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] !== quote) continue;
    if (text[i + 1] !== quote) return i + 1;
    i++;
  }
  return text.length;
}

/** The index after a `[…]` table or workbook reference (nested; `'` escapes the next character). */
function bracketEnd(text: string, start: number): number {
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    if (text[i] === "'") i++;
    else if (text[i] === '[') depth++;
    else if (text[i] === ']' && --depth === 0) return i + 1;
  }
  return text.length;
}

/**
 * A formula moved `rows` down and `columns` right, as Excel adjusts it on fill or copy: relative
 * A1 references shift, `$`-anchored parts stay. Whole-column (`F:F`) and whole-row (`1:1`)
 * references move only along their own axis. Text in double quotes, quoted sheet names
 * (`'Q1'!A1`), function names (`LOG10(`) and bracketed table references (`Table1[@[Q1]]`) are
 * left alone. A reference moved off the sheet becomes `#REF!`, as in Excel. A value that is not a
 * formula is returned as is.
 */
export function shiftFormula(formula: string, rows: number, columns: number): string {
  if (!formula.startsWith('=')) return formula;
  const row = (anchor: string, n: number) => (anchor ? n : n + rows);
  const column = (anchor: string, n: number) => (anchor ? n : n + columns);
  const rowOk = (n: number) => n >= 1 && n <= MAX_ROW;
  const columnOk = (n: number) => n >= 1 && n <= MAX_COLUMN;
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i]!;
    if (ch === '"' || ch === "'" || ch === '[') {
      const end = ch === '[' ? bracketEnd(formula, i) : quotedEnd(formula, i);
      out += formula.slice(i, end);
      i = end;
      continue;
    }
    const cell = matchAt(CELL_REF, formula, i);
    if (cell && columnOk(columnIndex(cell[2]!)) && rowOk(Number(cell[4]))) {
      const [text, ca = '', letters = '', ra = '', digits = ''] = cell;
      const c = column(ca, columnIndex(letters));
      const r = row(ra, Number(digits));
      out += columnOk(c) && rowOk(r) ? `${ca}${columnLabel(c)}${ra}${r}` : '#REF!';
      i += text.length;
      continue;
    }
    const cols = matchAt(COLUMN_SPAN, formula, i);
    if (cols && columnOk(columnIndex(cols[2]!)) && columnOk(columnIndex(cols[4]!))) {
      const [text, a1 = '', from = '', a2 = '', to = ''] = cols;
      const c1 = column(a1, columnIndex(from));
      const c2 = column(a2, columnIndex(to));
      out +=
        columnOk(c1) && columnOk(c2) ? `${a1}${columnLabel(c1)}:${a2}${columnLabel(c2)}` : '#REF!';
      i += text.length;
      continue;
    }
    const spanned = matchAt(ROW_SPAN, formula, i);
    if (spanned && rowOk(Number(spanned[2])) && rowOk(Number(spanned[4]))) {
      const [text, a1 = '', from = '', a2 = '', to = ''] = spanned;
      const r1 = row(a1, Number(from));
      const r2 = row(a2, Number(to));
      out += rowOk(r1) && rowOk(r2) ? `${a1}${r1}:${a2}${r2}` : '#REF!';
      i += text.length;
      continue;
    }
    const word = matchAt(WORD, formula, i);
    const next = word ? i + word[0].length : i + 1;
    out += formula.slice(i, next);
    i = next;
  }
  return out;
}

/**
 * The grid for writing one value over a whole range (`set I2:I57 =IF(A2=A1,…)`), like Excel's
 * Ctrl+Enter: the formula as written goes to the first cell and shifts for every other cell; a
 * plain value repeats. Undefined for a single cell, so `set` keeps writing just that cell. Throws
 * for a reversed range, a range over the cell budget, or a fill that moves a reference off the sheet.
 */
export function fillRange(address: string, value: string): string[][] | undefined {
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*):\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(
    address.slice(address.lastIndexOf('!') + 1),
  );
  if (!match) return undefined;
  const rows = Number(match[4]) - Number(match[2]) + 1;
  const columns = columnIndex(match[3]!) - columnIndex(match[1]!) + 1;
  if (rows < 1 || columns < 1)
    throw new Error('The range must go from its top-left cell to its bottom-right cell.');
  if (rows * columns === 1) return undefined;
  if (rows * columns > 100_000) throw new Error('Write dimensions exceed the cell budget.');
  const cells = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => shiftFormula(value, r, c)),
  );
  // Excel would write #REF! here; a fill the model chose should not break the formula, so refuse.
  // Counted, so a `#REF!` already in the formula (even inside quoted text) cannot hide a new one.
  const refs = (text: string) => text.split('#REF!').length - 1;
  if (cells.some((row) => row.some((v) => refs(v) > refs(value))))
    throw new Error('The fill moves a reference off the worksheet; choose a smaller range.');
  return cells;
}
