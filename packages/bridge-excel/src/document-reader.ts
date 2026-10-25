import { nativeRun } from './native-run.js';
import { documentOffset, type DocumentPage, type DocumentReadOptions } from '@ge/contracts';
import { isSet } from './capabilities-runtime.js';

export const DOCUMENT_TILE_CELLS = 1024;
export const DOCUMENT_TILE_COLUMNS = 32;
export const DOCUMENT_TILE_ROWS = 64;

/** All worksheets, including hidden sheets; only bounded used-range rectangles load values. */
export async function readExcelPage(options: DocumentReadOptions): Promise<DocumentPage> {
  options.signal?.throwIfAborted();
  if (!isSet('ExcelApi', '1.4'))
    throw new Error('ExcelApi 1.4 is required for paged workbook reads.');
  const cursor = options.cursor?.split(':') ?? ['0', '0', '0'];
  if (cursor.length !== 3) throw new Error('Invalid worksheet cursor.');
  const sheetIndex = documentOffset(cursor[0]);
  const row = documentOffset(cursor[1]);
  const col = documentOffset(cursor[2]);
  return nativeRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/id,items/name,items/visibility');
    await ctx.sync();
    const sheet = sheets.items[sheetIndex];
    const limitations = [
      'Cell values and formulas only; charts, images, comments and formatting are not extracted. Hidden worksheets are included.',
    ];
    if (!sheet) {
      if (sheetIndex || row || col) throw new Error('The workbook changed. Start a new read.');
      return { scope: 'Excel workbook cells', units: [], limitations };
    }
    const used = sheet.getUsedRangeOrNullObject(true);
    used.load('address,rowCount,columnCount,rowIndex,columnIndex,isNullObject');
    await ctx.sync();
    options.signal?.throwIfAborted();
    const nextSheet = sheetIndex + 1 < sheets.items.length ? `${sheetIndex + 1}:0:0` : undefined;
    if (used.isNullObject) {
      if (row || col) throw new Error('The worksheet changed. Start a new read.');
      return {
        scope: 'Excel workbook cells',
        limitations,
        units: [
          {
            id: `sheet:${sheet.id}`,
            text: JSON.stringify({ name: sheet.name, visibility: sheet.visibility, empty: true }),
          },
        ],
        ...(nextSheet ? { nextCursor: nextSheet } : {}),
      };
    }
    if (row >= used.rowCount || col >= used.columnCount)
      throw new Error('The used range changed. Start a new read.');
    const width = Math.min(DOCUMENT_TILE_COLUMNS, used.columnCount - col);
    // Fixed row stride across column tiles, including the narrow final tile.
    const height = Math.min(
      DOCUMENT_TILE_ROWS,
      Math.floor(DOCUMENT_TILE_CELLS / Math.min(DOCUMENT_TILE_COLUMNS, used.columnCount)),
      used.rowCount - row,
    );
    const range = sheet.getRangeByIndexes(
      used.rowIndex + row,
      used.columnIndex + col,
      height,
      width,
    );
    range.load('address,values,formulas');
    await ctx.sync();
    options.signal?.throwIfAborted();
    const units: DocumentPage['units'] = range.values.map((values, index) => {
      const formulas = Object.fromEntries(
        (range.formulas[index] ?? []).flatMap((value, column) =>
          typeof value === 'string' && value.startsWith('=')
            ? [[String(used.columnIndex + col + column + 1), value]]
            : [],
        ),
      );
      return {
        id: `sheet:${sheet.id}/row:${used.rowIndex + row + index + 1}/col:${used.columnIndex + col + 1}`,
        parentId: `sheet:${sheet.id}/tile:${row}:${col}`,
        text: JSON.stringify({ values, ...(Object.keys(formulas).length ? { formulas } : {}) }),
      };
    });
    units.unshift({
      id: `sheet:${sheet.id}/tile:${row}:${col}`,
      text: JSON.stringify({
        name: sheet.name,
        visibility: sheet.visibility,
        usedRange: used.address,
        rows: used.rowCount,
        columns: used.columnCount,
        tile: range.address,
      }),
    });
    const nextCursor =
      col + width < used.columnCount
        ? `${sheetIndex}:${row}:${col + width}`
        : row + height < used.rowCount
          ? `${sheetIndex}:${row + height}:0`
          : nextSheet;
    return {
      scope: 'Excel workbook cells',
      units,
      limitations,
      ...(nextCursor ? { nextCursor } : {}),
    };
  });
}
