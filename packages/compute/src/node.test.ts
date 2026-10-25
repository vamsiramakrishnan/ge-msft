import { describe, expect, it } from 'vitest';
import { ArtifactStore } from './artifacts.js';
import { createNodeCompute, NodeCompute } from '@ge/compute/node';

async function artifact() {
  return new ArtifactStore().add({
    title: 'Exact amounts',
    labels: ['amount'],
    rows: [['9007199254740993.01'], ['0.09']],
    sources: [],
    lineage: { parents: [], operation: 'snapshot' },
  });
}
describe('Node-only real DuckDB analytics', () => {
  it('preserves exact decimal results and bounds output', async () => {
    const table = await artifact();
    const compute = createNodeCompute({ maxRows: 1 });
    try {
      const total = await compute.query(
        `SELECT SUM(CAST(c0 AS DECIMAL(38,2))) AS total FROM ${table.id}`,
        [table],
      );
      expect(total.rows).toEqual([['9007199254740993.10']]);
      const bounded = await compute.query(`SELECT c0 FROM ${table.id}`, [table]);
      expect(bounded.rows).toHaveLength(1);
      expect(bounded.truncated).toBe(true);
    } finally {
      compute.dispose();
    }
  }, 30000);
  it('rejects mutated artifacts, unadmitted prior tables and external reads', async () => {
    const table = await artifact();
    const compute = createNodeCompute();
    try {
      await expect(
        compute.query(`SELECT * FROM ${table.id}`, [{ ...table, rows: [['999']] }]),
      ).rejects.toThrow('integrity');
      await compute.query(`SELECT * FROM ${table.id}`, [table]);
      await expect(compute.query(`SELECT * FROM ${table.id}`, [])).rejects.toThrow();
      await expect(
        compute.query("SELECT * FROM read_csv('https://example.test/data')", []),
      ).rejects.toThrow('unavailable');
    } finally {
      compute.dispose();
    }
  }, 30000);
  it('honors pre-execution cancellation and makes disposal terminal', async () => {
    const compute = createNodeCompute();
    const controller = new AbortController();
    controller.abort();
    await expect(compute.query('SELECT 1', [], controller.signal)).rejects.toMatchObject({
      name: 'AbortError',
    });
    compute.dispose();
    compute.dispose();
    await expect(compute.query('SELECT 1', [])).rejects.toThrow('closed');
  });
  it('refuses conflicting concurrent initialization and disposal during startup', async () => {
    const compute = createNodeCompute();
    const first = compute.query('SELECT 1', []);
    await expect(compute.query('SELECT 2', [])).rejects.toThrow('already running');
    compute.dispose();
    await expect(first).rejects.toThrow('closed');
  }, 30000);
  it('rejects invalid resource settings rather than silently weakening them', () => {
    expect(() => new NodeCompute({ maxRows: 10001 })).toThrow('maxRows');
    expect(() => new NodeCompute({ maxQueryMs: Infinity })).toThrow('maxQueryMs');
  });
  it('terminates an actually running expensive aggregate at its deadline and recovers with a fresh worker', async () => {
    const table = await new ArtifactStore().add({
      title: 'Deadline fixture',
      labels: ['n'],
      rows: Array.from({ length: 10000 }, (_, n) => [n]),
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    const limits = { maxQueryMs: 250 };
    const compute = createNodeCompute(limits);
    try {
      await expect(compute.query('SELECT 1 AS n', [])).resolves.toMatchObject({ rows: [[1]] });
      limits.maxQueryMs = 60000;
      const started = performance.now();
      await expect(
        compute.query(
          `SELECT SUM(ABS(a.c0-b.c0+c.c0)) AS total FROM ${table.id} a CROSS JOIN ${table.id} b CROSS JOIN ${table.id} c`,
          [table],
        ),
      ).rejects.toThrow('time budget');
      expect(performance.now() - started).toBeLessThan(5000);
      await expect(compute.query('SELECT 2 AS n', [])).resolves.toMatchObject({ rows: [[2]] });
      await expect(compute.query(`SELECT * FROM ${table.id}`, [])).rejects.toThrow();
    } finally {
      compute.dispose();
    }
  }, 30000);
  it('cancels synchronous WASM execution from the parent without blocking its event loop', async () => {
    const table = await new ArtifactStore().add({
      title: 'Cancellation fixture',
      labels: ['n'],
      rows: Array.from({ length: 10000 }, (_, n) => [n]),
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    const compute = createNodeCompute();
    try {
      await compute.query('SELECT 1 AS n', []);
      const controller = new AbortController();
      const started = performance.now();
      const timer = setTimeout(() => controller.abort(), 150);
      try {
        await expect(
          compute.query(
            `SELECT SUM(ABS(a.c0-b.c0+c.c0)) AS total FROM ${table.id} a CROSS JOIN ${table.id} b CROSS JOIN ${table.id} c`,
            [table],
            controller.signal,
          ),
        ).rejects.toMatchObject({ name: 'AbortError' });
      } finally {
        clearTimeout(timer);
      }
      expect(performance.now() - started).toBeLessThan(5000);
      await expect(compute.query('SELECT 3 AS n', [])).resolves.toMatchObject({ rows: [[3]] });
    } finally {
      compute.dispose();
    }
  }, 30000);
});
