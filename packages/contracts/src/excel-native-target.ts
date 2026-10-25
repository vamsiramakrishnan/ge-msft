function hasControl(value: string): boolean {
  return [...value].some((character) => character.charCodeAt(0) < 32);
}
function invalidSheetCharacters(value: string): boolean {
  return (
    hasControl(value) ||
    ['[', ']', ':', '*', '?', '/', '\\'].some((character) => value.includes(character))
  );
}

/** Explicit finite rectangles only: no active-sheet, names, external books, or whole columns. */
export function boundedExcelRange(selector: string):
  | {
      sheet: string;
      address: string;
      rows: number;
      columns: number;
      firstRow: number;
      firstColumn: number;
    }
  | undefined {
  const match =
    /^(?:'((?:[^']|'')+)'|([^!'[\]:]+))!\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$/.exec(
      selector,
    );
  if (!match) return undefined;
  const sheet = match[1]?.replace(/''/g, "'") ?? match[2]!;
  if (!sheet.trim() || sheet.length > 31 || invalidSheetCharacters(sheet)) return undefined;
  const col = (s: string) =>
    [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const firstColumn = col(match[3]!);
  const lastColumn = col(match[5] ?? match[3]!);
  const firstRow = Number(match[4]);
  const lastRow = Number(match[6] ?? match[4]);
  const rows = lastRow - firstRow + 1;
  const columns = lastColumn - firstColumn + 1;
  if (
    rows < 1 ||
    columns < 1 ||
    lastRow > 1_048_576 ||
    lastColumn > 16_384 ||
    columns > 256 ||
    rows * columns > 10_000
  )
    return undefined;
  return { sheet, address: selector, rows, columns, firstRow, firstColumn };
}
