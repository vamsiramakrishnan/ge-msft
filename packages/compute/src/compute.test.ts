import { beforeAll, afterAll, describe, expect, it } from 'vitest';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import type * as DuckDB from '@duckdb/duckdb-wasm/blocking';
import { ArtifactStore } from './artifacts.js';
import { validateQuery } from './sql-policy.js';
import { reconciliationQuery } from './reconcile.js';
import { exactDecimalColumnSql } from './exact-decimal.js';
import { artifactToIPC, arrowRows, ENGINE_SETTINGS } from './arrow.js';
import { makeCellSnapshot } from '@ge/contracts';

const require = createRequire(import.meta.url);
const duck = require('@duckdb/duckdb-wasm/blocking') as typeof DuckDB;
let db: Awaited<ReturnType<typeof duck.createDuckDB>>;
let conn: ReturnType<typeof db.connect>;
beforeAll(async () => {
  const dist = resolve(require.resolve('@duckdb/duckdb-wasm'), '..');
  db = await duck.createDuckDB(
    {
      mvp: {
        mainModule: resolve(dist, 'duckdb-mvp.wasm'),
        mainWorker: resolve(dist, 'duckdb-node-mvp.worker.cjs'),
      },
    },
    new duck.VoidLogger(),
    duck.NODE_RUNTIME,
  );
  await db.instantiate();
  db.open({ query: { castBigIntToDouble: false, castDecimalToDouble: false } });
  conn = db.connect();
  conn.query(ENGINE_SETTINGS);
}, 30000);
afterAll(() => {
  conn?.close();
  db?.reset();
});

describe('actual DuckDB WASM computation', () => {
  it('reconciles exact decimals, multiple payments, currencies, invalid and unmatched rows', async () => {
    const store = new ArtifactStore();
    const left = await store.add({
      title: 'Invoices',
      labels: ['key', 'amount', 'currency'],
      rows: [
        ['A', '0.30', 'USD'],
        ['B', '9007199254740993.01', 'USD'],
        ['C', '7.00', 'EUR'],
        ['D', 'oops', 'USD'],
        ['E', '1.00', 'BAD!'],
      ],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    const right = await store.add({
      title: 'Payments',
      labels: ['key', 'amount', 'currency'],
      rows: [
        ['A', '0.10', 'USD'],
        ['A', '0.20', 'USD'],
        ['B', '9007199254740993.00', 'USD'],
        ['C', '7.00', 'USD'],
        ['Z', '2.00', 'USD'],
      ],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    conn.insertArrowFromIPCStream(artifactToIPC(left), { name: left.id, create: true });
    conn.insertArrowFromIPCStream(artifactToIPC(right), { name: right.id, create: true });
    const sql = validateQuery(
      reconciliationQuery(left, right, {
        left: left.id,
        right: right.id,
        leftKey: 0,
        rightKey: 0,
        leftAmount: 1,
        rightAmount: 1,
        leftCurrency: 2,
        rightCurrency: 2,
        tolerance: '0.001',
      }),
    );
    expect(conn.getTableNames(sql).sort()).toEqual([left.id, right.id].sort());
    const rows = arrowRows(conn.query(sql), 100);
    expect(rows.find((r) => r[0] === 'A')).toEqual([
      'A',
      'USD',
      '0.300000',
      '0.300000',
      '0.000000',
      'matched',
      '1',
      '2',
    ]);
    expect(rows.find((r) => r[0] === 'B')?.slice(2, 6)).toEqual([
      '9007199254740993.010000',
      '9007199254740993.000000',
      '0.010000',
      'variance',
    ]);
    expect(
      rows
        .filter((r) => r[0] === 'C')
        .map((r) => r[5])
        .sort(),
    ).toEqual(['unallocated', 'unpaid']);
    expect(rows.filter((r) => r[5] === 'invalid')).toHaveLength(2);
    expect(rows.find((r) => r[0] === 'Z')?.[5]).toBe('unallocated');
  });
  it('admits a header-only table and computes a zero count', async () => {
    const store = new ArtifactStore();
    const empty = await store.add({
      title: 'Empty',
      labels: ['id'],
      rows: [],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    conn.insertArrowFromIPCStream(artifactToIPC(empty), { name: empty.id, create: true });
    expect(arrowRows(conn.query(`SELECT count(*) FROM ${empty.id}`), 10)).toEqual([['0']]);
  });
  it('never rounds invalid reconciliation decimals or presents partial group totals', async () => {
    const store = new ArtifactStore();
    const left = await store.add({
      title: 'Precision invoices',
      labels: ['key', 'amount'],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
      rows: [
        ['fine', '0.0000001'],
        ['exponent', '1e2'],
        ['overflow', '100000000000000000000000000000000'],
        ['partial', '7.00'],
        ['partial', 'not a number'],
        ['blank', null],
        ['exact', '9007199254740993.010001'],
        ['negative', '-0.000001'],
      ],
    });
    const right = await store.add({
      title: 'Precision payments',
      labels: ['key', 'amount'],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
      rows: [
        ['fine', '0.00'],
        ['exponent', '100.00'],
        ['overflow', '0.00'],
        ['partial', '7.00'],
        ['blank', '0.00'],
        ['exact', '9007199254740993.010001'],
        ['negative', '-0.000001'],
      ],
    });
    conn.insertArrowFromIPCStream(artifactToIPC(left), { name: left.id, create: true });
    conn.insertArrowFromIPCStream(artifactToIPC(right), { name: right.id, create: true });
    const sql = validateQuery(
      reconciliationQuery(left, right, {
        left: left.id,
        right: right.id,
        leftKey: 0,
        rightKey: 0,
        leftAmount: 1,
        rightAmount: 1,
        currency: 'USD',
        tolerance: '0',
      }),
    );
    const rows = arrowRows(conn.query(sql), 100);
    for (const key of ['fine', 'exponent', 'overflow', 'partial', 'blank'])
      expect(rows.find((row) => row[0] === key)?.slice(2, 6)).toEqual([
        null,
        null,
        null,
        'invalid',
      ]);
    expect(rows.find((row) => row[0] === 'exact')?.slice(2, 6)).toEqual([
      '9007199254740993.010001',
      '9007199254740993.010001',
      '0.000000',
      'matched',
    ]);
    expect(rows.find((row) => row[0] === 'negative')?.slice(2, 6)).toEqual([
      '-0.000001',
      '-0.000001',
      '0.000000',
      'matched',
    ]);
  });
  it('rejects already-imprecise native numbers on either reconciliation side before SQL execution', async () => {
    const store = new ArtifactStore();
    const unsafe = await store.add({
      title: 'Unsafe amounts',
      labels: ['key', 'amount'],
      rows: [['A', Number.MAX_SAFE_INTEGER + 1]],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    const exact = await store.add({
      title: 'Exact amounts',
      labels: ['key', 'amount'],
      rows: [['A', '9007199254740992']],
      sources: [],
      lineage: { parents: [], operation: 'snapshot' },
    });
    for (const [left, right] of [
      [unsafe, exact],
      [exact, unsafe],
    ]) {
      expect(() =>
        reconciliationQuery(left!, right!, {
          left: left!.id,
          right: right!.id,
          leftKey: 0,
          rightKey: 0,
          leftAmount: 1,
          rightAmount: 1,
          currency: 'USD',
          tolerance: '0',
        }),
      ).toThrow('Store those amounts as decimal text');
    }
    expect(() =>
      reconciliationQuery(exact, exact, {
        left: exact.id,
        right: exact.id,
        leftKey: 5,
        rightKey: 0,
        leftAmount: 1,
        rightAmount: 1,
        currency: 'USD',
        tolerance: '0',
      }),
    ).toThrow('Column 6 (index 5)');
    expect(() => exactDecimalColumnSql('c0); SELECT 1')).toThrow('captured column identifier');
  });
  it('roundtrips decimals and integers without Number coercion', () => {
    expect(
      arrowRows(
        conn.query('SELECT CAST(-0.000001 AS DECIMAL(38,6)), CAST(9007199254740993 AS BIGINT)'),
        10,
      ),
    ).toEqual([['-0.000001', '9007199254740993']]);
  });
  it('disables external access and locks security settings in the real engine', () => {
    expect(() => conn.query("SELECT * FROM read_csv('https://example.com/source.csv')")).toThrow();
    expect(() => conn.query('SET enable_external_access=true')).toThrow();
    expect(() => validateQuery('INSTALL httpfs')).toThrow();
  });
});

describe('admission and artifact ownership', () => {
  it.each([
    "SELECT * FROM read_csv('x')",
    'SELECT 1; DROP TABLE x',
    'WITH RECURSIVE x AS (SELECT 1) SELECT * FROM x',
    "SELECT * FROM query_table('x')",
    "SELECT current_setting('secret')",
    'SELECT 1 -- comment',
    'PRAGMA version',
    "SELECT * FROM glob('*')",
    'SELECT random()',
  ])('rejects %s before dispatch', (sql) => expect(() => validateQuery(sql)).toThrow());
  it('keeps SQL-looking text as data and artifacts immutable and addressable', async () => {
    const store = new ArtifactStore();
    const snap = await makeCellSnapshot({
      surface: 'excel',
      documentId: 'd',
      objectId: 's',
      locator: 'S!A1:B2',
      values: [
        ['id', 'amount'],
        ['=WEBSERVICE("bad")', 2],
      ],
    });
    const a = await store.fromSnapshot(snap);
    const b = await store.fromSnapshot({ ...snap, capturedAt: '2020-01-01T00:00:00.000Z' });
    expect(a.id).toBe(b.id);
    a.rows[0]![1] = 99;
    expect(store.get(a.id).rows[0]![1]).toBe(2);
    const list = store.list();
    list[0]!.preview[0]![1] = 300;
    expect(store.get(a.id).rows[0]![1]).toBe(2);
    store.remove(a.id);
    expect(() => store.get(a.id)).toThrow('no longer');
  });
});
