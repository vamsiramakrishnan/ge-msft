import { describe, expect, it, vi } from 'vitest';
import { makeCellSnapshot, type CellValue } from '@ge/contracts';
import { AnalysisWorkspace } from './analysis-workspace.js';
import type { DocBridge } from './bridge.js';
function fixture() {
  let values: CellValue[][] = [
    ['id', 'amount'],
    ['A', '0.30'],
  ];
  const bridge = {
    surface: 'excel',
    captureCells: async (locator: string) =>
      makeCellSnapshot({ surface: 'excel', documentId: 'd', locator, values }),
  } as DocBridge;
  const query = vi.fn(async () => ({
    columns: ['item_key', 'status'],
    rows: [['A', 'variance']],
    truncated: false,
    durationMs: 1,
  }));
  const workspace = new AnalysisWorkspace(bridge, async () => ({ query, dispose() {} }));
  return {
    workspace,
    query,
    edit() {
      values = [
        ['id', 'amount'],
        ['A', '0.31'],
      ];
    },
  };
}
it('checks source freshness before compute, inspection, filtering and materialization', async () => {
  const f = fixture();
  const a = await f.workspace.execute({ kind: 'capture', range: 'S!A1:B2' });
  f.edit();
  await expect(
    f.workspace.execute({ kind: 'query', sql: `SELECT * FROM ${a!.id}`, inputs: [a!.id] }),
  ).rejects.toThrow('Source changed');
  await expect(f.workspace.materialize(a!.id, 'S!D1')).rejects.toThrow('Source changed');
  await expect(f.workspace.execute({ kind: 'inspect', id: a!.id })).rejects.toThrow(
    'Source changed',
  );
  expect(f.query).not.toHaveBeenCalled();
});
it('preserves truncation through derived results and never writes a partial artifact', async () => {
  const f = fixture();
  const source = await f.workspace.execute({ kind: 'capture', range: 'S!A1:B2' });
  f.query.mockResolvedValueOnce({
    columns: ['item_key', 'status'],
    rows: [['A', 'variance']],
    truncated: true,
    durationMs: 1,
  });
  const result = await f.workspace.execute({
    kind: 'query',
    sql: `SELECT * FROM ${source!.id}`,
    inputs: [source!.id],
  });
  const child = await f.workspace.execute({
    kind: 'query',
    sql: `SELECT * FROM ${result!.id}`,
    inputs: [result!.id],
  });
  expect(child?.truncated).toBe(true);
  await expect(f.workspace.materialize(child!.id, 'S!D1')).rejects.toThrow('truncated');
});
it('rejects stale column selections with an actionable error before dispatching SQL', async () => {
  const f = fixture();
  const source = await f.workspace.execute({ kind: 'capture', range: 'S!A1:B2' });
  await expect(
    f.workspace.execute({
      kind: 'query',
      inputs: [source!.id],
      sql: `SELECT c5 FROM ${source!.id}`,
      requiredColumns: [{ input: source!.id, indices: [5] }],
    }),
  ).rejects.toThrow('Column 6 (index 5)');
  await expect(
    f.workspace.execute({
      kind: 'query',
      inputs: [source!.id],
      sql: `SELECT * FROM ${source!.id}`,
      requiredColumns: [{ input: 'a_unknown', indices: [0] }],
    }),
  ).rejects.toThrow('undeclared query input');
  expect(f.query).not.toHaveBeenCalled();
});
it('rejects unsafe native numeric amounts before claiming exact decimal arithmetic', async () => {
  const f = fixture();
  const source = await f.workspace.artifacts.add({
    title: 'Unsafe amounts',
    labels: ['amount'],
    rows: [[Number.MAX_SAFE_INTEGER + 1]],
    sources: [],
    lineage: { parents: [], operation: 'snapshot' },
  });
  await expect(
    f.workspace.execute({
      kind: 'query',
      inputs: [source.id],
      sql: `SELECT c0 FROM ${source.id}`,
      requiredColumns: [{ input: source.id, indices: [0], exactDecimal: true }],
    }),
  ).rejects.toThrow('Store those amounts as decimal text');
  expect(f.query).not.toHaveBeenCalled();
});
describe('typed action offers', () => {
  it('derives counts from findings and preserves meaningful labels when filtering', async () => {
    const f = fixture();
    const a = await f.workspace.execute({ kind: 'capture', range: 'S!A1:B2' });
    const b = await f.workspace.execute({ kind: 'capture', range: 'S!D1:E2' });
    await f.workspace.execute({
      kind: 'reconcile',
      spec: {
        left: a!.id,
        right: b!.id,
        leftKey: 0,
        rightKey: 0,
        leftAmount: 1,
        rightAmount: 1,
        currency: 'USD',
      },
    });
    expect(f.workspace.state().offers).toMatchObject([
      { title: '1 variance', count: 1, approval: 'none' },
    ]);
    const offer = f.workspace.state().offers[0]!;
    f.query.mockResolvedValueOnce({
      columns: ['c0', 'c1'],
      rows: [['A', 'variance']],
      truncated: false,
      durationMs: 1,
    });
    const filtered = await f.workspace.execute(offer.action);
    expect(filtered?.columns.map((c) => c.label)).toEqual(['item_key', 'status']);
    f.workspace.artifacts.remove(offer.artifactId);
    await expect(f.workspace.execute(offer.action)).rejects.toThrow('no longer');
  });
});
