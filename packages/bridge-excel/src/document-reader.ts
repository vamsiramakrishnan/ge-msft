import { nativeRun } from './native-run.js';
import { documentOffset, type DocumentPage, type DocumentReadOptions } from '@ge/contracts';
import { isSet } from './capabilities-runtime.js';
import { excelRowUnit, formulaDependencies, type FormulaAlias } from './formula-dependencies.js';

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
    // Count before admitting native collections. Large alias catalogs remain explicit gaps;
    // they never cause an unbounded named-range or table value read.
    const names = isSet('ExcelApi', '1.7') ? ctx.workbook.names : undefined;
    const tables = ctx.workbook.tables;
    const nameCount = typeof names?.getCount === 'function' ? names.getCount() : undefined;
    const tableCount = typeof tables?.getCount === 'function' ? tables.getCount() : undefined;
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
    const localNames = isSet('ExcelApi', '1.7') ? sheet.names : undefined;
    const localNameCount =
      typeof localNames?.getCount === 'function' ? localNames.getCount() : undefined;
    const loadNames = nameCount !== undefined && boundedAliasCount(nameCount.value);
    const loadTables = tableCount !== undefined && boundedAliasCount(tableCount.value);
    if (loadNames) names!.load('items/name,items/type,items/formula');
    if (loadTables) tables.load('items/name');
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
            relationsComplete: true,
            text: JSON.stringify({ name: sheet.name, visibility: sheet.visibility, empty: true }),
          },
        ],
        ...(nextSheet ? { nextCursor: nextSheet } : {}),
      };
    }
    if (row >= used.rowCount || col >= used.columnCount)
      throw new Error('The used range changed. Start a new read.');
    const width = Math.min(
      DOCUMENT_TILE_COLUMNS - ((used.columnIndex + col) % DOCUMENT_TILE_COLUMNS),
      used.columnCount - col,
    );
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
    const loadLocalNames = localNameCount !== undefined && boundedAliasCount(localNameCount.value);
    if (loadLocalNames) localNames!.load('items/name,items/type,items/formula');
    if (loadNames && names!.items.length !== nameCount!.value)
      throw new Error('Named range metadata changed. Start a new read.');
    if (loadTables && tables.items.length !== tableCount!.value)
      throw new Error('Table metadata changed. Start a new read.');
    const tableRanges = loadTables ? tables.items.map((table) => table.getRange()) : [];
    for (const tableRange of tableRanges) tableRange.load('address');
    await ctx.sync();
    options.signal?.throwIfAborted();
    if (loadLocalNames && localNames!.items.length !== localNameCount!.value)
      throw new Error('Worksheet name metadata changed. Start a new read.');
    const capturedAliases: FormulaAlias[] = [
      ...(loadLocalNames
        ? localNames!.items
            .filter((name) => name.type === 'Range')
            .map((name) => ({
              name: `${sheet.name}!${name.name.slice(name.name.lastIndexOf('!') + 1)}`,
              address: String(name.formula),
              kind: 'named-range' as const,
            }))
        : []),
      // A local definition shadows a workbook name. Unknown local coverage cannot authorize
      // resolution of an unqualified global alias, even if a same-named global item exists.
      ...(loadNames && loadLocalNames
        ? names!.items
            .filter((name) => name.type === 'Range')
            .map((name) => ({
              name: name.name,
              address: String(name.formula),
              kind: 'named-range' as const,
            }))
        : []),
      ...(loadTables
        ? tables.items.map((table, index) => ({
            name: table.name,
            address: tableRanges[index]!.address,
            kind: 'table' as const,
          }))
        : []),
    ];
    const aliases: FormulaAlias[] = [];
    let aliasBytes = 0;
    for (const alias of capturedAliases) {
      const bytes = new TextEncoder().encode(JSON.stringify(alias)).byteLength;
      if (
        aliases.length >= 256 ||
        alias.name.length > 512 ||
        alias.address.length > 4096 ||
        aliasBytes + bytes > 16 * 1024
      )
        continue;
      aliases.push(alias);
      aliasBytes += bytes;
    }
    let incompleteDependencies = false;
    const units: DocumentPage['units'] = range.values.map((values, index) => {
      const formulas = Object.fromEntries(
        (range.formulas[index] ?? []).flatMap((value, column) =>
          typeof value === 'string' && value.startsWith('=')
            ? [[String(used.columnIndex + col + column + 1), value]]
            : [],
        ),
      );
      const related = new Set<string>();
      let relationsComplete = true;
      for (const [column, formula] of Object.entries(formulas)) {
        // Range.formulas also returns literal values. Ambiguous formula-looking text cannot
        // establish a dependency. The capture still retains the original native text as data.
        if (values[Number(column) - used.columnIndex - col - 1] === formula) {
          incompleteDependencies = true;
          relationsComplete = false;
          continue;
        }
        const dependencies = formulaDependencies(
          formula as string,
          sheet,
          sheets.items,
          100 - related.size,
          aliases,
        );
        for (const id of dependencies.relatedIds) related.add(id);
        if (!dependencies.complete) {
          incompleteDependencies = true;
          relationsComplete = false;
        }
      }
      return {
        id: excelRowUnit(sheet.id, used.rowIndex + row + index + 1, used.columnIndex + col + 1),
        parentId: `sheet:${sheet.id}/tile:${row}:${col}`,
        relationsComplete,
        ...(related.size ? { relatedIds: [...related] } : {}),
        text: JSON.stringify({ values, ...(Object.keys(formulas).length ? { formulas } : {}) }),
      };
    });
    if (incompleteDependencies)
      limitations.push(
        'Some formula dependencies are unresolved or exceed the 100-links-per-row budget. Dynamic, external, whole-row/column, 3D and unresolved named/structured references require explicit reads; formula-looking literal text grants no links.',
      );
    units.unshift({
      id: `sheet:${sheet.id}/tile:${row}:${col}`,
      relationsComplete: true,
      text: JSON.stringify({
        name: sheet.name,
        visibility: sheet.visibility,
        usedRange: used.address,
        rows: used.rowCount,
        columns: used.columnCount,
        tile: range.address,
        // Bind alias changes to source freshness even when they resolve to empty cells.
        aliases,
        omittedAliases:
          !loadNames || !loadLocalNames || !loadTables || aliases.length < capturedAliases.length,
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

function boundedAliasCount(value: number): boolean {
  if (!Number.isSafeInteger(value) || value < 0) throw new Error('Invalid native alias count.');
  return value <= 128;
}
