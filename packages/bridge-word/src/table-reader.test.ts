import { afterEach, describe, expect, it, vi } from 'vitest';
import { OfficeWordHost } from './host-port.js';
import { WordBridge } from './word-bridge.js';
import { readWordTableTile } from './table-reader.js';
import { AssistSession } from '@ge/runtime';

afterEach(() => {
  vi.unstubAllGlobals();
});

function install(grid: string[][], options: { extraRows?: boolean; changed?: boolean } = {}) {
  const cellReads: Array<[number, number]> = [];
  const tableLoads: unknown[] = [];
  const rowLoads: unknown[] = [];
  let tableWindow = { skip: 0, top: 0 };
  let rowWindow = { skip: 0, top: 0 };
  let syncs = 0;
  const table = {
    get rowCount() {
      return grid.length + (options.changed && syncs >= 3 ? 1 : 0);
    },
    load: (value: unknown) => tableLoads.push(value),
    rows: {
      load: (value: { skip: number; top: number }) => {
        rowWindow = value;
        rowLoads.push(value);
      },
      get items() {
        return grid
          .slice(rowWindow.skip, rowWindow.skip + rowWindow.top + (options.extraRows ? 1 : 0))
          .map((row) => ({ cellCount: row.length }));
      },
    },
    getCell: (r: number, c: number) => {
      cellReads.push([r, c]);
      return { value: grid[r]![c], load: (value: string) => expect(value).toBe('value') };
    },
  };
  vi.stubGlobal('Office', {
    context: { document: { url: 'word://saved' }, requirements: { isSetSupported: () => true } },
  });
  vi.stubGlobal('Word', {
    run: async (work: (ctx: unknown) => Promise<unknown>) =>
      work({
        document: {
          body: {
            tables: {
              load: (value: { skip: number; top: number }) => {
                tableWindow = value;
                tableLoads.push(value);
              },
              get items() {
                return tableWindow.skip === 0 ? [table] : [];
              },
            },
          },
        },
        sync: async () => {
          syncs++;
        },
      }),
  });
  return { cellReads, tableLoads, rowLoads };
}

describe('bounded native Word structural table reads', () => {
  it('reads at most 20 by 16 actual cells without whole-table values and gives exact continuation selectors', async () => {
    const grid = Array.from({ length: 50 }, (_, r) =>
      Array.from({ length: 30 }, (_, c) => `${r}:${c}`),
    );
    const seen = install(grid);
    const first = await readWordTableTile(1, 0, 0);
    expect(seen.cellReads).toHaveLength(320);
    expect(first.nextSelector).toBe('table:1:0:16');
    expect(first.complete).toBe(false);
    expect(first.rowCount).toBe(50);
    expect(seen.tableLoads[0]).toEqual({ select: 'rowCount', skip: 0, top: 1 });
    expect(seen.rowLoads[0]).toEqual({ select: 'cellCount', skip: 0, top: 20 });
    const second = await readWordTableTile(1, 0, 16);
    expect(second.nextSelector).toBe('table:1:20:0');
    expect(second.rows[0]!.cells[0]).toEqual({ columnIndex: 16, text: '0:16' });
  });
  it('preserves ragged row metadata and refuses missing or changed table windows', async () => {
    install([['a'], ['b', 'c']]);
    expect((await readWordTableTile(1, 0, 0)).rows.map((row) => row.cellCount)).toEqual([1, 2]);
    await expect(readWordTableTile(2, 0, 0)).rejects.toThrow('missing');
    await expect(readWordTableTile(1, 4, 0)).rejects.toThrow('out of range');
    install([['a']], { changed: true });
    await expect(readWordTableTile(1, 0, 0)).rejects.toThrow('changed during capture');
  });
  it('refuses a host which ignores the row budget', async () => {
    install(
      Array.from({ length: 30 }, () => ['a']),
      { extraRows: true },
    );
    await expect(readWordTableTile(1, 0, 0)).rejects.toThrow('bounded table-row');
  });
  it('uses byte budgets without splitting Unicode and reports every omitted text byte', async () => {
    install([['😀'.repeat(10000), 'kept']]);
    const tile = await readWordTableTile(1, 0, 0);
    expect(new TextEncoder().encode(tile.rows[0]!.cells[0]!.text).byteLength).toBe(2048);
    expect(tile.rows[0]!.cells[0]!.text).not.toContain('�');
    expect(tile.omittedTextBytes).toBe(40000 - 2048);
    expect(tile.complete).toBe(false);
  });
  it('integrates the existing runtime readRange boundary and preserves limitations as source data', async () => {
    install([['Amount', '120.00']]);
    const result = await new WordBridge(new OfficeWordHost()).readRange('table:1');
    expect(result[0]!.ref).toMatchObject({ kind: 'table', surface: 'word' });
    expect(result[0]!.value.as).toBe('text');
    if (result[0]!.value.as === 'text') {
      const tile = JSON.parse(result[0]!.value.text);
      expect(tile.rows[0].cells[1].text).toBe('120.00');
      expect(tile.limitations.join(' ')).toContain('does not authorize a write');
    }
    expect(await new WordBridge().readRange('anything-else')).toEqual([]);
    await expect(new WordBridge().readRange('table:invalid')).rejects.toThrow('structural reads');
  });

  it('keeps inspect phrase on the actual runtime literal-search fallback', async () => {
    class SearchBridge extends WordBridge {
      override async listContext() {
        return [];
      }
      override async searchDocument(query: string) {
        return [
          {
            ref: {
              id: 'match:1',
              kind: 'paragraph' as const,
              surface: 'word' as const,
              title: query,
            },
            value: { as: 'text' as const, text: 'Needle phrase from native search' },
          },
        ];
      }
    }
    const bridge = new SearchBridge();
    const search = vi.spyOn(bridge, 'searchDocument');
    const session = new AssistSession(
      bridge,
      {} as ConstructorParameters<typeof AssistSession>[1],
      {
        unit: { connectors: [], surfaceContext: { kind: 'word' } },
        context: { docState: false, lazyRead: false },
        primeOnHostEvent: false,
      },
    );
    const events: unknown[] = [];
    for await (const event of session.runCommandProgram('inspect "Needle phrase"\ndone'))
      events.push(event);
    expect(search).toHaveBeenCalledWith('Needle phrase');
    expect(JSON.stringify(events)).toContain('Needle phrase from native search');
  });
});

describe('Word inserted-table postconditions', () => {
  it.each(['verified', 'mismatch', 'unknown'] as const)(
    'preserves the landed insert and records %s readback',
    async (status) => {
      let syncs = 0;
      const expected = [['Amount', '120.00']];
      const table = {
        rowCount: 1,
        values: status === 'mismatch' ? [['Amount', 'wrong']] : expected,
        load: vi.fn(),
      };
      vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } });
      vi.stubGlobal('Word', {
        InsertLocation: { after: 'After' },
        run: async (work: (ctx: unknown) => Promise<unknown>) =>
          work({
            document: { getSelection: () => ({ insertTable: () => table }) },
            sync: async () => {
              syncs++;
              if (status === 'unknown' && syncs === 2) throw new Error('readback failed');
            },
          }),
      });
      const result = await new OfficeWordHost({ trackEdits: false }).insertTable(
        undefined,
        { matchCase: false },
        1,
        2,
        expected,
        () => 0,
      );
      expect(result).toMatchObject({
        status: 'applied',
        location: 'selection',
        verification: { status },
      });
      expect(table.load).toHaveBeenCalledWith('values,rowCount');
    },
  );
});
