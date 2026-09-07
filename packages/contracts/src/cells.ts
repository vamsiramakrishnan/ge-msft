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
