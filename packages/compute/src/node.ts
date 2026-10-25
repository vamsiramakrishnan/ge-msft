import { Worker } from 'node:worker_threads';
import { URL as NodeURL } from 'node:url';
import { TableArtifactSchema, contentHash, type TableArtifact } from '@ge/contracts';
import { validateQuery } from './sql-policy.js';
import type { NodeComputeJob } from './node-worker-protocol.js';
import type { ComputeEngine, QueryResult } from './types.js';

export interface NodeComputeOptions {
  maxRows?: number;
  /** Parent-enforced deadline after the isolated worker is ready; expiry terminates that worker. */
  maxQueryMs?: number;
}
const aborted = (): DOMException => new DOMException('Analysis cancelled', 'AbortError');
const record = (raw: unknown): Record<string, unknown> | undefined =>
  raw !== null && typeof raw === 'object' && !Array.isArray(raw)
    ? (raw as Record<string, unknown>)
    : undefined;

function result(raw: unknown, limit: number): QueryResult {
  const value = record(raw);
  if (!value || !Array.isArray(value.columns) || !Array.isArray(value.rows))
    throw new Error('Compute worker returned an invalid result.');
  const columns = value.columns;
  const rows = value.rows;
  if (
    columns.length < 1 ||
    columns.length > 256 ||
    !columns.every((column) => typeof column === 'string' && column.length <= 32768) ||
    rows.length > limit ||
    rows.length * columns.length > 100000 ||
    typeof value.truncated !== 'boolean' ||
    !rows.every(
      (row) =>
        Array.isArray(row) &&
        row.length === columns.length &&
        row.every(
          (cell) =>
            cell === null ||
            typeof cell === 'boolean' ||
            (typeof cell === 'string' && cell.length <= 32768) ||
            (typeof cell === 'number' && Number.isFinite(cell)),
        ),
    ) ||
    new TextEncoder().encode(JSON.stringify(value)).byteLength > 8 * 1024 * 1024 + 65536
  )
    throw new Error('Compute worker returned an invalid or oversized result.');
  return {
    columns: columns as string[],
    rows: rows as QueryResult['rows'],
    truncated: value.truncated,
    durationMs: 0,
  };
}

/** Node analytics run in a terminable worker; no synchronous model-authored SQL executes on the caller thread. */
export class NodeCompute implements ComputeEngine {
  private worker?: Worker;
  private busy = false;
  private disposed = false;
  private cancelActive?: () => void;
  private readonly options: Readonly<NodeComputeOptions>;
  constructor(options: NodeComputeOptions = {}) {
    const maxRows = options.maxRows;
    const maxQueryMs = options.maxQueryMs;
    this.options = Object.freeze({ maxRows, maxQueryMs });
    if (maxRows !== undefined && (!Number.isInteger(maxRows) || maxRows < 1 || maxRows > 10000))
      throw new Error('Node compute maxRows must be between 1 and 10000.');
    if (
      maxQueryMs !== undefined &&
      (!Number.isFinite(maxQueryMs) || maxQueryMs < 1 || maxQueryMs > 60000)
    )
      throw new Error('Node compute maxQueryMs must be between 1 and 60000.');
  }
  private assertOpen(signal?: AbortSignal): void {
    if (this.disposed) throw new Error('The compute workspace is closed.');
    signal?.throwIfAborted();
  }
  private async terminate(worker: Worker): Promise<void> {
    if (this.worker === worker) this.worker = undefined;
    try {
      await worker.terminate();
    } catch {
      // Do not create replacement CPU work if this worker's termination could not be confirmed.
      this.disposed = true;
    }
  }
  private initialize(signal?: AbortSignal): Promise<Worker> {
    this.assertOpen(signal);
    if (this.worker) return Promise.resolve(this.worker);
    const worker = new Worker(new NodeURL('./node-worker.js', import.meta.url));
    worker.unref();
    this.worker = worker;
    // Active operations also reject; this listener handles idle EventEmitter errors.
    worker.on('error', () => {
      if (this.worker === worker) this.worker = undefined;
    });
    worker.on('exit', () => {
      if (this.worker === worker) this.worker = undefined;
    });
    return new Promise((resolve, reject) => {
      let settled = false;
      const deadline = performance.now() + 30000;
      const cleanup = () => {
        clearTimeout(timer);
        signal?.removeEventListener('abort', cancel);
        worker.off('message', message);
        worker.off('error', failure);
        worker.off('exit', exit);
        if (this.cancelActive === cancel) this.cancelActive = undefined;
      };
      const stop = (error: Error) => {
        if (settled) return;
        settled = true;
        cleanup();
        void this.terminate(worker).then(() => reject(error));
      };
      const cancel = () =>
        stop(this.disposed ? new Error('The compute workspace is closed.') : aborted());
      const failure = () => stop(new Error('The isolated compute worker could not start.'));
      const exit = () => stop(new Error('The isolated compute worker stopped during startup.'));
      const message = (raw: unknown) => {
        if (performance.now() > deadline) {
          stop(new Error('Compute startup exceeded its 30-second time budget.'));
          return;
        }
        if (record(raw)?.type !== 'ready') {
          failure();
          return;
        }
        if (settled) return;
        settled = true;
        cleanup();
        resolve(worker);
      };
      const timer = setTimeout(
        () => stop(new Error('Compute startup exceeded its 30-second time budget.')),
        30000,
      );
      this.cancelActive = cancel;
      worker.on('message', message);
      worker.on('error', failure);
      worker.on('exit', exit);
      signal?.addEventListener('abort', cancel, { once: true });
      if (signal?.aborted || this.disposed) cancel();
    });
  }
  private execute(worker: Worker, job: NodeComputeJob, signal?: AbortSignal): Promise<QueryResult> {
    return new Promise((resolve, reject) => {
      let settled = false;
      const deadline = performance.now() + (this.options.maxQueryMs ?? 10000);
      const cleanup = () => {
        clearTimeout(timer);
        signal?.removeEventListener('abort', cancel);
        worker.off('message', message);
        worker.off('error', failure);
        worker.off('exit', exit);
        if (this.cancelActive === cancel) this.cancelActive = undefined;
      };
      const stop = (error: Error) => {
        if (settled) return;
        settled = true;
        cleanup();
        // Publish rejection after termination, so a new query cannot race the old CPU task.
        void this.terminate(worker).then(() => reject(error));
      };
      const cancel = () =>
        stop(this.disposed ? new Error('The compute workspace is closed.') : aborted());
      const failure = () => stop(new Error('The isolated compute worker failed.'));
      const exit = () =>
        stop(new Error('The isolated compute worker stopped before a result was received.'));
      const message = (raw: unknown) => {
        if (performance.now() > deadline) {
          stop(new Error('Analysis exceeded its time budget. Narrow the query.'));
          return;
        }
        const value = record(raw);
        if (value?.id !== job.id) {
          stop(new Error('The compute receipt does not match the requested query.'));
          return;
        }
        if (value.type !== 'result') {
          stop(
            new Error(
              'The analytical query could not complete within its admitted inputs and limits.',
            ),
          );
          return;
        }
        try {
          const output = result(value.result, job.maxRows);
          if (performance.now() > deadline) {
            stop(new Error('Analysis exceeded its time budget. Narrow the query.'));
            return;
          }
          if (settled) return;
          settled = true;
          cleanup();
          resolve(output);
        } catch {
          stop(new Error('Compute worker returned an invalid or oversized result.'));
        }
      };
      const timer = setTimeout(
        () => stop(new Error('Analysis exceeded its time budget. Narrow the query.')),
        this.options.maxQueryMs ?? 10000,
      );
      this.cancelActive = cancel;
      worker.on('message', message);
      worker.on('error', failure);
      worker.on('exit', exit);
      signal?.addEventListener('abort', cancel, { once: true });
      if (signal?.aborted || this.disposed) {
        cancel();
        return;
      }
      try {
        worker.postMessage(job);
      } catch {
        failure();
      }
    });
  }
  async query(
    raw: string,
    tables: readonly TableArtifact[],
    signal?: AbortSignal,
  ): Promise<QueryResult> {
    this.assertOpen(signal);
    if (this.busy) throw new Error('An analysis is already running.');
    const sql = validateQuery(raw);
    if (tables.length > 32) throw new Error('Too many input tables.');
    const input = tables.map((table) => TableArtifactSchema.parse(table));
    if (new TextEncoder().encode(JSON.stringify(input)).byteLength > 32 * 1024 * 1024)
      throw new Error('Compute input exceeds 32 MiB.');
    this.busy = true;
    const started = performance.now();
    try {
      for (const artifact of input) {
        const { columns, rows, sources, lineage, truncated } = artifact;
        const hash = await contentHash({ columns, rows, sources, lineage, truncated });
        if (hash !== artifact.hash || artifact.id !== `a_${hash.slice(7, 31)}`)
          throw new Error('Artifact integrity check failed. Capture the source again.');
      }
      this.assertOpen(signal);
      const worker = await this.initialize(signal);
      this.assertOpen(signal);
      const output = await this.execute(
        worker,
        {
          version: 1,
          id: crypto.randomUUID(),
          sql,
          tables: input,
          maxRows: this.options.maxRows ?? 5000,
        },
        signal,
      );
      this.assertOpen(signal);
      return { ...output, durationMs: performance.now() - started };
    } finally {
      this.busy = false;
    }
  }
  dispose(): void {
    this.disposed = true;
    if (this.cancelActive) this.cancelActive();
    else if (this.worker) void this.terminate(this.worker);
  }
}
export function createNodeCompute(options: NodeComputeOptions = {}): ComputeEngine {
  return new NodeCompute(options);
}
