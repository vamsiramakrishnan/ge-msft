/**
 * A tiny, self-contained A1-notation helper for the Excel simulator. Just enough to translate
 * between `"A1"` / `"A1:D9"` references and zero-based `{ row, col }` spans, plus column letters.
 * Kept dependency-free and pure so the fake's address arithmetic is itself unit-testable.
 */

/** A zero-based rectangular span resolved from an A1 reference. */
export interface A1Span {
  startRow: number;
  startCol: number;
  rows: number;
  cols: number;
}

/** A 2-D string grid (row-major). */
export type Grid = string[][];

/** Column letters → zero-based index: `A`→0, `Z`→25, `AA`→26. */
export function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Zero-based index → column letters: 0→`A`, 26→`AA`. */
export function indexToCol(index: number): string {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Split a single A1 cell like `"B3"` into `{ row, col }` (both zero-based). */
function parseCell(cell: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(cell.trim());
  if (!m) throw new Error(`fake-excel: not an A1 cell: "${cell}"`);
  return { col: colToIndex(m[1] as string), row: Number.parseInt(m[2] as string, 10) - 1 };
}

/** Parse `"A1"` or `"A1:D9"` (sheet prefix already stripped) into a zero-based {@link A1Span}. */
export function parseA1(ref: string): A1Span {
  const [a, b] = ref.split(':');
  const start = parseCell(a as string);
  if (!b) return { startRow: start.row, startCol: start.col, rows: 1, cols: 1 };
  const end = parseCell(b);
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    rows: Math.abs(end.row - start.row) + 1,
    cols: Math.abs(end.col - start.col) + 1,
  };
}

/** Build a sheet-qualified address from a top-left A1 cell + a row/col span. */
export function addressOf(sheet: string, topLeft: string, rows: number, cols: number): string {
  const { row, col } = parseCell(topLeft);
  if (rows <= 1 && cols <= 1) return `${sheet}!${cellRef(col, row)}`;
  const end = cellRef(col + cols - 1, row + rows - 1);
  return `${sheet}!${cellRef(col, row)}:${end}`;
}

/** A zero-based column+row → an A1 cell ref (`col=1,row=2` → `"B3"`). */
export function cellRef(col: number, row: number): string {
  return `${indexToCol(col)}${row + 1}`;
}
