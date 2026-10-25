import type { ActuationKind, Surface } from '@ge/contracts';

export interface NativeTrace {
  id: string;
  surface: Surface;
  stage: 'dispatch' | 'batch';
  kind?: ActuationKind;
  changeId?: string;
  syncs: number;
  syncMs: number;
  durationMs: number;
  outcome: 'completed' | 'refused' | 'unknown';
}
/** Fixed metadata only: never record arguments, document identity, errors, body or query text. */
export class NativeTraceStore {
  private readonly entries: NativeTrace[] = [];
  record(trace: NativeTrace): void {
    this.entries.push({
      id: trace.id,
      surface: trace.surface,
      stage: trace.stage,
      ...(trace.kind ? { kind: trace.kind } : {}),
      ...(trace.changeId ? { changeId: trace.changeId } : {}),
      syncs: trace.syncs,
      syncMs: trace.syncMs,
      durationMs: trace.durationMs,
      outcome: trace.outcome,
    });
    if (this.entries.length > 200) this.entries.shift();
  }
  list(): NativeTrace[] {
    return structuredClone(this.entries);
  }
  clear(): void {
    this.entries.length = 0;
  }
}
export const nativeTraces = new NativeTraceStore();
/** Instrument request-local sync only, restoring it before Office releases the context. */
export async function observeNativeBatch<C extends { sync(): Promise<unknown> }, T>(
  surface: Surface,
  context: C,
  work: (context: C) => Promise<T>,
  store = nativeTraces,
): Promise<T> {
  const started = performance.now();
  const original = context.sync;
  const trace: NativeTrace = {
    id: crypto.randomUUID(),
    surface,
    stage: 'batch',
    syncs: 0,
    syncMs: 0,
    durationMs: 0,
    outcome: 'unknown',
  };
  context.sync = async function () {
    trace.syncs++;
    const start = performance.now();
    try {
      return await original.call(context);
    } finally {
      trace.syncMs += performance.now() - start;
    }
  };
  try {
    const result = await work(context);
    trace.outcome = 'completed';
    return result;
  } finally {
    context.sync = original;
    trace.durationMs = performance.now() - started;
    store.record(trace);
  }
}
