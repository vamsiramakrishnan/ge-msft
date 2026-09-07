import { describe, expect, it } from 'vitest';
import { snapshotRange } from './cell-snapshot.js';
const context = { sync: async () => {} } as Excel.RequestContext;
function range(formulaCell: boolean): Excel.Range {
  return {
    address: 'Sheet1!A1',
    rowCount: 1,
    columnCount: 1,
    isNullObject: false,
    values: [['=literal']],
    formulas: [['=literal']],
    load() {},
    worksheet: { id: 'sheet-1', load() {} },
    getSpecialCellsOrNullObject: () => ({
      isNullObject: !formulaCell,
      load() {},
      areas: { items: [{ address: 'Sheet1!A1' }], load() {} },
    }),
  } as unknown as Excel.Range;
}
describe('native snapshot semantics', () => {
  it('does not promote formula-looking literal text into an undo formula', async () => {
    const result = await snapshotRange(range(false), context, 'doc');
    expect(result.values).toEqual([['=literal']]);
    expect(result.formulas).toEqual([['']]);
  });
  it('preserves an actual formula even if it evaluates to its own expression text', async () => {
    const result = await snapshotRange(range(true), context, 'doc');
    expect(result.formulas).toEqual([['=literal']]);
  });
  it('refuses an ambiguous snapshot on a host without native formula areas', async () => {
    const r = range(false);
    (r as unknown as { getSpecialCellsOrNullObject?: unknown }).getSpecialCellsOrNullObject =
      undefined;
    await expect(snapshotRange(r, context, 'doc')).rejects.toThrow('distinguish');
  });
  it('includes worksheet identity in the version hash', async () => {
    const first = await snapshotRange(range(false), context, 'doc');
    const secondRange = range(false);
    Object.defineProperty(secondRange.worksheet, 'id', { value: 'replacement-sheet' });
    const second = await snapshotRange(secondRange, context, 'doc');
    expect(first.hash).not.toBe(second.hash);
  });
});
