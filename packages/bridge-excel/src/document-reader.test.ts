import { afterEach, describe, expect, it, vi } from 'vitest';
import { DOCUMENT_TILE_CELLS, readExcelPage } from './document-reader.js';

afterEach(() => {
  vi.unstubAllGlobals();
});
describe('Excel paged reads', () => {
  it('indexes bounded native names/tables in the same three sync stages without loading their values', async () => {
    let syncs = 0;
    const namesLoad = vi.fn();
    const tablesLoad = vi.fn();
    const tableRange = {
      address: 'Sales!A2:B3',
      load: vi.fn((properties: string) => expect(properties).toBe('address')),
      get values(): never {
        throw new Error('Table values must not be read for dependencies');
      },
    };
    const names = {
      getCount: () => ({ value: 1 }),
      load: namesLoad,
      items: [{ name: 'Revenue', type: 'Range', formula: '=Sales!A2' }],
    };
    const tables = {
      getCount: () => ({ value: 1 }),
      load: tablesLoad,
      items: [{ name: 'SalesTable', getRange: () => tableRange }],
    };
    const sheet = {
      id: 'native-sales',
      name: 'Sales',
      visibility: 'Visible',
      names: { getCount: () => ({ value: 0 }), load: () => {}, items: [] },
      getUsedRangeOrNullObject: () => ({
        address: 'Sales!A1:B3',
        rowCount: 3,
        columnCount: 2,
        rowIndex: 0,
        columnIndex: 0,
        isNullObject: false,
        load: () => {},
      }),
      getRangeByIndexes: () => ({
        address: 'Sales!A1:B3',
        load: () => {},
        values: [
          [100, 200],
          [1, 2],
          [3, 4],
        ],
        formulas: [
          ['=Revenue', '=SUM(SalesTable[Amount])'],
          ['', ''],
          ['', ''],
        ],
      }),
    };
    vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
    vi.stubGlobal('Excel', {
      run: async (work: (context: unknown) => Promise<unknown>) =>
        work({
          workbook: { worksheets: { load: () => {}, items: [sheet] }, names, tables },
          sync: async () => {
            syncs++;
          },
        }),
    });
    const page = await readExcelPage({});
    expect(syncs).toBe(3);
    expect(page.units[1]).toMatchObject({
      relationsComplete: true,
      relatedIds: ['sheet:native-sales/row:2/col:1', 'sheet:native-sales/row:3/col:1'],
    });
    expect(JSON.parse(page.units[0]!.text)).toMatchObject({
      omittedAliases: false,
      aliases: [
        { name: 'Revenue', address: '=Sales!A2', kind: 'named-range' },
        { name: 'SalesTable', address: 'Sales!A2:B3', kind: 'table' },
      ],
    });
    expect(namesLoad).toHaveBeenCalledTimes(1);
    expect(tablesLoad).toHaveBeenCalledTimes(1);
    const localNames = {
      getCount: () => ({ value: 1 }),
      load: () => {},
      items: [{ name: 'Revenue', type: 'Range', formula: '=Sales!A10' }],
    };
    Object.assign(sheet, { names: localNames });
    const shadowed = await readExcelPage({});
    expect(shadowed.units[1]!.relatedIds).toContain('sheet:native-sales/row:10/col:1');
    localNames.getCount = () => ({ value: 129 });
    const unknownLocals = await readExcelPage({});
    expect(unknownLocals.units[1]!.relationsComplete).toBe(false);
    expect(
      JSON.parse(unknownLocals.units[0]!.text).aliases.some(
        (alias: { name: string }) => alias.name === 'Revenue',
      ),
    ).toBe(false);
    // Native counts prevent even metadata collection loads above the explicit cap.
    names.getCount = () => ({ value: 129 });
    tables.getCount = () => ({ value: 10000 });
    namesLoad.mockClear();
    tablesLoad.mockClear();
    const limited = await readExcelPage({});
    expect(namesLoad).not.toHaveBeenCalled();
    expect(tablesLoad).not.toHaveBeenCalled();
    expect(limited.units[1]!.relationsComplete).toBe(false);
    expect(JSON.parse(limited.units[0]!.text).omittedAliases).toBe(true);
  });
  it.each([8, 70])(
    'covers 30 tabs × 1500 rows × %i columns without a full used-range value load',
    async (columns) => {
      let cells = 0;
      const rectangles = new Set<string>();
      vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
      vi.stubGlobal('Excel', {
        run: async (fn: (ctx: unknown) => Promise<unknown>) =>
          fn({
            workbook: {
              worksheets: {
                load: () => undefined,
                items: Array.from({ length: 30 }, (_, index) => ({
                  id: `s${index}`,
                  name: `Tab ${index}`,
                  visibility: index === 29 ? 'Hidden' : 'Visible',
                  getUsedRangeOrNullObject: () => ({
                    address: `'Tab ${index}'!C5:ZZ1504`,
                    rowCount: 1500,
                    columnCount: columns,
                    rowIndex: 4,
                    columnIndex: 2,
                    isNullObject: false,
                    load: (props: string) => {
                      expect(props).not.toContain('values');
                    },
                    get values(): never {
                      throw new Error('Unbounded values load');
                    },
                  }),
                  getRangeByIndexes: (row: number, col: number, height: number, width: number) => {
                    expect(height * width).toBeLessThanOrEqual(DOCUMENT_TILE_CELLS);
                    const key = `${index}:${row}:${col}`;
                    expect(rectangles.has(key)).toBe(false);
                    rectangles.add(key);
                    cells += height * width;
                    return {
                      address: `${key}:${height}:${width}`,
                      load: () => undefined,
                      values: Array.from({ length: height }, (_, i) =>
                        Array.from({ length: width }, (_, j) => (row + i) * columns + col + j),
                      ),
                      formulas: Array.from({ length: height }, () =>
                        Array.from({ length: width }, () => ''),
                      ),
                    };
                  },
                })),
              },
            },
            sync: async () => undefined,
          }),
      });
      let cursor: string | undefined;
      const unitIds = new Set<string>();
      let last = '';
      do {
        const page = await readExcelPage({ cursor });
        for (const unit of page.units) {
          expect(unitIds.has(unit.id)).toBe(false);
          unitIds.add(unit.id);
          last = unit.id;
        }
        cursor = page.nextCursor;
      } while (cursor);
      expect(cells).toBe(30 * 1500 * columns);
      expect(last).toContain('sheet:s29/row:1504/');
    },
    20_000,
  );
});
