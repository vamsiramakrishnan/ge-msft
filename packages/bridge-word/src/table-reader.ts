import { nativeRun } from './native-run.js';
import { isSet } from './capabilities-runtime.js';

export interface WordTableTile {
  tableIndex: number;
  rowCount: number;
  rowOffset: number;
  columnOffset: number;
  rows: Array<{
    rowIndex: number;
    cellCount: number;
    cells: Array<{ columnIndex: number; text: string }>;
  }>;
  omittedTextBytes: number;
  /** True only when this single capture contains every table cell and all of its text. */
  complete: boolean;
  nextSelector?: string;
}

export const WORD_TABLE_LIMITS = Object.freeze({
  rows: 20,
  columns: 16,
  textBytes: 24 * 1024,
  cellBytes: 2048,
});

/** Native row/cell windows; no load of Table.values or the full table's cell collection. */
export async function readWordTableTile(
  tableIndex: number,
  rowOffset: number,
  columnOffset: number,
): Promise<WordTableTile> {
  if (
    ![tableIndex, rowOffset, columnOffset].every(Number.isSafeInteger) ||
    tableIndex < 1 ||
    tableIndex > 10000 ||
    rowOffset < 0 ||
    columnOffset < 0
  )
    throw new Error('Invalid Word table window.');
  if (!isSet('WordApi', '1.3')) throw new Error('Native table reads require WordApi 1.3.');
  return nativeRun(async (ctx) => {
    const tables = ctx.document.body.tables;
    tables.load({ select: 'rowCount', skip: tableIndex - 1, top: 1 });
    await ctx.sync();
    if (tables.items.length !== 1)
      throw new Error('The table is missing or Word did not honor the table window.');
    const table = tables.items[0]!;
    if (!Number.isSafeInteger(table.rowCount) || table.rowCount < 1 || rowOffset >= table.rowCount)
      throw new Error('The table changed or the row offset is out of range.');
    const rowCount = table.rowCount;
    const rows = table.rows;
    rows.load({ select: 'cellCount', skip: rowOffset, top: WORD_TABLE_LIMITS.rows });
    await ctx.sync();
    if (rows.items.length !== Math.min(WORD_TABLE_LIMITS.rows, rowCount - rowOffset))
      throw new Error('Word did not honor bounded table-row reads.');
    const cells = rows.items.map((row, i) => {
      if (!Number.isSafeInteger(row.cellCount) || row.cellCount < 1)
        throw new Error('The table cell structure could not be read.');
      return Array.from(
        { length: Math.max(0, Math.min(WORD_TABLE_LIMITS.columns, row.cellCount - columnOffset)) },
        (_, j) => {
          const cell = table.getCell(rowOffset + i, columnOffset + j);
          cell.load('value');
          return cell;
        },
      );
    });
    table.load('rowCount');
    await ctx.sync();
    if (table.rowCount !== rowCount) throw new Error('The table changed during capture.');
    let budget = WORD_TABLE_LIMITS.textBytes;
    let omittedTextBytes = 0;
    const resultRows = rows.items.map((row, i) => ({
      rowIndex: rowOffset + i,
      cellCount: row.cellCount,
      cells: cells[i]!.map((cell, j) => {
        if (typeof cell.value !== 'string') throw new Error('The table cell could not be read.');
        const raw = new TextEncoder().encode(cell.value);
        let end = Math.min(raw.byteLength, budget, WORD_TABLE_LIMITS.cellBytes);
        let text = '';
        while (end > 0) {
          try {
            text = new TextDecoder('utf-8', { fatal: true }).decode(raw.subarray(0, end));
            break;
          } catch {
            end--;
          }
        }
        budget -= end;
        omittedTextBytes += raw.byteLength - end;
        return { columnIndex: columnOffset + j, text };
      }),
    }));
    const moreColumns = rows.items.some(
      (row) => row.cellCount > columnOffset + WORD_TABLE_LIMITS.columns,
    );
    if (columnOffset > 0 && resultRows.every((row) => row.cells.length === 0))
      throw new Error('The table changed or the column offset is out of range.');
    const nextSelector = moreColumns
      ? `table:${tableIndex}:${rowOffset}:${columnOffset + WORD_TABLE_LIMITS.columns}`
      : rowOffset + rows.items.length < rowCount
        ? `table:${tableIndex}:${rowOffset + rows.items.length}:0`
        : undefined;
    return {
      tableIndex,
      rowCount,
      rowOffset,
      columnOffset,
      rows: resultRows,
      omittedTextBytes,
      complete:
        rowOffset === 0 &&
        columnOffset === 0 &&
        nextSelector === undefined &&
        omittedTextBytes === 0,
      ...(nextSelector ? { nextSelector } : {}),
    };
  });
}
