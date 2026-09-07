import * as duckdb from '@duckdb/duckdb-wasm';
import { artifactToIPC, arrowRows, ENGINE_SETTINGS } from './arrow.js';
import { TableArtifactSchema, contentHash, type TableArtifact } from '@ge/contracts';
import { validateQuery } from './sql-policy.js';
import type { ComputeEngine, QueryResult } from './types.js';

export interface BrowserComputeOptions {
  workerUrl: string;
  wasmUrl: string;
  timeoutMs?: number;
  maxRows?: number;
}
/** No Office/model/auth dependency. The only caller-supplied code is the restricted SQL dialect. */
export class BrowserCompute implements ComputeEngine {
  private worker?: Worker;
  private database?: duckdb.AsyncDuckDB;
  private connection?: duckdb.AsyncDuckDBConnection;
  private seeded = new Set<string>();
  private busy = false;
  private disposed = false;
  private generation = 0;
  private abortActive?: () => void;
  constructor(private readonly options: BrowserComputeOptions) {}

  async query(
    raw: string,
    tables: readonly TableArtifact[],
    signal?: AbortSignal,
  ): Promise<QueryResult> {
    if (this.disposed) throw new Error('The compute workspace is closed.');
    if (this.busy)
      throw new Error('An analysis is already running. Cancel it or wait for it to finish.');
    const sql = validateQuery(raw);
    if (tables.length > 32) throw new Error('Too many input tables.');
    const input = tables.map((t) => TableArtifactSchema.parse(t));
    if (new TextEncoder().encode(JSON.stringify(input)).byteLength > 32 * 1024 * 1024)
      throw new Error('Compute input exceeds 32 MiB.');
    signal?.throwIfAborted();
    this.busy = true;
    const started = performance.now();
    const epoch = this.generation;
    const controller = new AbortController();
    let timer: ReturnType<typeof setTimeout> | undefined;
    let cancel!: () => void;
    let timeout!: () => void;
    const cancelled = new Promise<never>((_, reject) => {
      cancel = () => {
        controller.abort();
        this.reset();
        reject(new DOMException('Analysis cancelled', 'AbortError'));
      };
      this.abortActive = cancel;
      signal?.addEventListener('abort', cancel, { once: true });
      timeout = () => {
        controller.abort();
        this.reset();
        reject(new Error('Analysis exceeded its time budget. Narrow the query and try again.'));
      };
      timer = setTimeout(timeout, this.connection ? (this.options.timeoutMs ?? 10_000) : 30_000);
    });
    try {
      return await Promise.race([
        cancelled,
        (async () => {
          for (const artifact of input) {
            const { columns, rows, sources, lineage, truncated } = artifact;
            const hash = await contentHash({ columns, rows, sources, lineage, truncated });
            if (hash !== artifact.hash || artifact.id !== `a_${hash.slice(7, 31)}`)
              throw new Error('Artifact integrity check failed. Capture the source again.');
          }
          controller.signal.throwIfAborted();
          const conn = await this.initialize(controller.signal, epoch);
          if (timer) clearTimeout(timer);
          timer = setTimeout(timeout, this.options.timeoutMs ?? 10_000);
          // Drop tables outside this request. Data from an older request cannot be queried by guessing its id.
          const admitted = new Set(input.map((t) => t.id));
          for (const id of this.seeded)
            if (!admitted.has(id)) {
              await conn.query(`DROP TABLE ${id}`);
              this.seeded.delete(id);
            }
          for (const artifact of input) {
            controller.signal.throwIfAborted();
            if (this.seeded.has(artifact.id)) continue;
            await conn.insertArrowFromIPCStream(artifactToIPC(artifact), {
              name: artifact.id,
              create: true,
            });
            this.seeded.add(artifact.id);
          }
          // Checking referenced tables closes paths such as FROM 'https://...' as well as access to
          // internal catalogs. Engine external access is independently disabled below.
          const names = await conn.getTableNames(sql);
          if (names.some((name) => !admitted.has(name)))
            throw new Error('Queries may only read the input artifacts listed in this workspace.');
          const limit = Math.max(1, Math.min(this.options.maxRows ?? 5000, 10_000));
          const result = await conn.query(
            `SELECT * FROM (${sql}) AS bounded_result LIMIT ${limit + 1}`,
          );
          const columns = result.schema.fields.map((field) => field.name);
          if (columns.length > 256) throw new Error('Query returned too many columns.');
          const rows = arrowRows(result, limit);
          controller.signal.throwIfAborted();
          return {
            columns,
            rows,
            truncated: result.numRows > limit,
            durationMs: performance.now() - started,
          };
        })(),
      ]);
    } catch (error) {
      this.reset();
      throw error;
    } finally {
      if (timer) clearTimeout(timer);
      signal?.removeEventListener('abort', cancel);
      this.abortActive = undefined;
      this.busy = false;
    }
  }
  private async initialize(
    signal: AbortSignal,
    epoch: number,
  ): Promise<duckdb.AsyncDuckDBConnection> {
    if (this.connection) return this.connection;
    const wasmUrl = new URL(this.options.wasmUrl, location.href);
    const workerUrl = new URL(this.options.workerUrl, location.href);
    if (wasmUrl.origin !== location.origin || workerUrl.origin !== location.origin)
      throw new Error('Compute assets must be served from the application origin.');
    const response = await fetch(wasmUrl, { signal, credentials: 'same-origin' });
    if (!response.ok)
      throw new Error('The compute engine could not load. Check the application deployment.');
    const blob = await response.blob();
    if (blob.size > 64 * 1024 * 1024) throw new Error('Compute binary exceeds the asset budget.');
    signal.throwIfAborted();
    if (epoch !== this.generation) throw new DOMException('Analysis cancelled', 'AbortError');
    const moduleUrl = URL.createObjectURL(new Blob([blob], { type: 'application/wasm' }));
    try {
      this.worker = new Worker(workerUrl);
      this.database = new duckdb.AsyncDuckDB(new duckdb.VoidLogger(), this.worker);
      await this.database.instantiate(moduleUrl);
      signal.throwIfAborted();
      await this.database.open({
        maximumThreads: 1,
        allowUnsignedExtensions: false,
        query: { castBigIntToDouble: false, castDecimalToDouble: false },
        opfs: { fileHandling: 'manual' },
      });
      const conn = await this.database.connect();
      // Fail startup if any required engine control is unavailable. No permissive fallback.
      await conn.query(ENGINE_SETTINGS);
      signal.throwIfAborted();
      this.connection = conn;
      return conn;
    } finally {
      URL.revokeObjectURL(moduleUrl);
    }
  }
  private reset(): void {
    this.generation++;
    this.worker?.terminate();
    this.database?.detach();
    this.worker = undefined;
    this.database = undefined;
    this.connection = undefined;
    this.seeded.clear();
  }
  dispose(): void {
    this.disposed = true;
    this.abortActive?.();
    this.reset();
  }
}
export function createBrowserCompute(options: BrowserComputeOptions): ComputeEngine {
  return new BrowserCompute(options);
}
