import { parentPort } from 'node:worker_threads';
import { createRequire } from 'node:module';
import { resolve } from 'node:path';
import type * as DuckDB from '@duckdb/duckdb-wasm/blocking';
import { TableArtifactSchema, contentHash } from '@ge/contracts';
import { artifactToIPC, arrowRows, ENGINE_SETTINGS } from './arrow.js';
import { validateQuery } from './sql-policy.js';
import type { NodeComputeJob } from './node-worker-protocol.js';

const port = parentPort;
if (!port) throw new Error('Node analytics worker needs a parent thread.');
const require = createRequire(import.meta.url);
const duck = require('@duckdb/duckdb-wasm/blocking') as typeof DuckDB;
const directory = resolve(require.resolve('@duckdb/duckdb-wasm'), '..');
let busy = false;
async function start(): Promise<void> {
  const database = await duck.createDuckDB(
    {
      mvp: {
        mainModule: resolve(directory, 'duckdb-mvp.wasm'),
        mainWorker: resolve(directory, 'duckdb-node-mvp.worker.cjs'),
      },
    },
    new duck.VoidLogger(),
    duck.NODE_RUNTIME,
  );
  await database.instantiate();
  database.open({ query: { castBigIntToDouble: false, castDecimalToDouble: false } });
  const connection = database.connect();
  connection.query(ENGINE_SETTINGS);
  const seeded = new Set<string>();
  port!.on('message', (raw: unknown) => {
    void (async () => {
      const input = raw as Partial<NodeComputeJob> | null;
      const id = typeof input?.id === 'string' ? input.id : 'invalid';
      if (busy) {
        port!.postMessage({ type: 'error', id });
        return;
      }
      busy = true;
      try {
        if (
          !input ||
          input.version !== 1 ||
          !/^[a-f0-9-]{36}$/.test(id) ||
          typeof input.sql !== 'string' ||
          !Array.isArray(input.tables) ||
          input.tables.length > 32 ||
          !Number.isInteger(input.maxRows) ||
          input.maxRows! < 1 ||
          input.maxRows! > 10000
        )
          throw new Error('Invalid compute message.');
        const sql = validateQuery(input.sql);
        const tables = input.tables.map((table) => TableArtifactSchema.parse(table));
        if (new TextEncoder().encode(JSON.stringify(tables)).byteLength > 32 * 1024 * 1024)
          throw new Error('Input budget exceeded.');
        for (const artifact of tables) {
          const { columns, rows, sources, lineage, truncated } = artifact;
          const hash = await contentHash({ columns, rows, sources, lineage, truncated });
          if (hash !== artifact.hash || artifact.id !== `a_${hash.slice(7, 31)}`)
            throw new Error('Invalid artifact.');
        }
        const admitted = new Set(tables.map((table) => table.id));
        for (const previous of seeded)
          if (!admitted.has(previous)) {
            connection.query(`DROP TABLE ${previous}`);
            seeded.delete(previous);
          }
        for (const artifact of tables)
          if (!seeded.has(artifact.id)) {
            connection.insertArrowFromIPCStream(artifactToIPC(artifact), {
              name: artifact.id,
              create: true,
            });
            seeded.add(artifact.id);
          }
        if (connection.getTableNames(sql).some((name) => !admitted.has(name)))
          throw new Error('Unadmitted table.');
        const result = connection.query(
          `SELECT * FROM (${sql}) AS bounded_result LIMIT ${input.maxRows! + 1}`,
        );
        port!.postMessage({
          type: 'result',
          id,
          result: {
            columns: result.schema.fields.map((field) => field.name),
            rows: arrowRows(result, input.maxRows!),
            truncated: result.numRows > input.maxRows!,
          },
        });
      } catch {
        port!.postMessage({ type: 'error', id });
      } finally {
        busy = false;
      }
    })();
  });
  port!.postMessage({ type: 'ready' });
}
void start().catch(() => {
  port.postMessage({ type: 'fatal' });
  port.close();
});
