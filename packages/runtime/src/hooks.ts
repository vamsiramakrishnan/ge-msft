import {
  ResolvedContextSchema,
  type ActuationRequest,
  type ActuationResult,
  type ResolvedContext,
  type SseEvent,
  type Surface,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import type { RunOutcome, TaskMode } from './execution-ledger.js';

/** Trusted, compiled extensions. Document text and model output can never register executable hooks. */
export interface RuntimeHookPayloads {
  'message:received': {
    mode: TaskMode;
    text: string;
    dataStoreSpecs?: Array<{ dataStore: string; filter?: string }>;
  };
  'model:request': { query: string; route: string };
  /** Delivered before the event reaches the panel or command parser. Earlier streamed tokens remain visible. */
  'model:event': { event: SseEvent };
  'model:response': { text: string; route: string };
  'tool:before': { name: string; args: unknown };
  'tool:after': { name: string; result: unknown };
  'plan:ready': { effects: ActuationRequest[] };
  'effect:before': { request: ActuationRequest };
  'effect:after': { request: ActuationRequest; result: ActuationResult };
  'task:verify': { outcome: RunOutcome };
  'task:finished': { outcome: RunOutcome };
  'host:event': { event: HostEvent };
}
export type RuntimeHookPhase = keyof RuntimeHookPayloads;
export type DeepReadonly<T> = T extends object
  ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;
export interface HookContext {
  readonly taskId: string;
  readonly surface: Surface;
  readonly signal: AbortSignal;
}
export type HookDecision = void | { kind: 'continue' } | { kind: 'block'; reason: string };
export type HookResult<K extends RuntimeHookPhase> =
  | HookDecision
  | (K extends 'message:received' ? { kind: 'context'; entries: ResolvedContext[] } : never);
export interface RuntimeHook<K extends RuntimeHookPhase = RuntimeHookPhase> {
  id: string;
  on: K;
  mode: 'guard' | 'observe';
  /** Higher priority runs first; registration order breaks ties. */
  priority?: number;
  /** 1–10,000 ms; default 750. Async cancellation is cooperative, not a JS sandbox. */
  timeoutMs?: number;
  handle(
    payload: DeepReadonly<RuntimeHookPayloads[K]>,
    context: HookContext,
  ): HookResult<K> | Promise<HookResult<K>>;
}
export interface HookRecord {
  sequence: number;
  taskId: string;
  hookId: string;
  phase: RuntimeHookPhase;
  outcome: 'continued' | 'blocked' | 'context' | 'error' | 'timeout' | 'cancelled';
  durationMs: number;
}
export class HookBlockedError extends Error {
  constructor(
    readonly hookId: string,
    readonly phase: RuntimeHookPhase,
    message: string,
  ) {
    super(message);
    this.name = 'HookBlockedError';
  }
}
const GUARD_PHASES = new Set<RuntimeHookPhase>([
  'message:received',
  'model:request',
  'model:event',
  'model:response',
  'tool:before',
  'plan:ready',
  'effect:before',
  'task:verify',
]);
const PHASES = new Set<RuntimeHookPhase>([
  ...GUARD_PHASES,
  'tool:after',
  'effect:after',
  'task:finished',
  'host:event',
]);
const MAX_HOOKS = 64;
const MAX_RECORDS = 256;
const MAX_CONTEXT_BYTES = 64 * 1024;

/** Deterministic hook dispatch with isolated snapshots and metadata-only, bounded diagnostics. */
export class RuntimeHooks {
  private readonly hooks: RuntimeHook[] = [];
  private readonly history: HookRecord[] = [];
  private readonly listeners = new Set<(record: HookRecord) => void>();
  private readonly running = new Set<string>();
  private sequence = 0;

  register<K extends RuntimeHookPhase>(hook: RuntimeHook<K>): () => void {
    if (
      !PHASES.has(hook.on) ||
      !['guard', 'observe'].includes(hook.mode) ||
      typeof hook.handle !== 'function'
    )
      throw new Error('Invalid runtime hook definition');
    if (!/^[a-zA-Z0-9][a-zA-Z0-9._:/-]{0,119}$/.test(hook.id)) throw new Error('Invalid hook id');
    if (this.hooks.some((h) => h.id === hook.id)) throw new Error(`Duplicate hook id: ${hook.id}`);
    if (this.hooks.length >= MAX_HOOKS) throw new Error('Runtime hook limit reached');
    if (hook.mode === 'guard' && !GUARD_PHASES.has(hook.on))
      throw new Error(`${hook.on} is observation-only`);
    if (
      hook.timeoutMs !== undefined &&
      (!Number.isFinite(hook.timeoutMs) || hook.timeoutMs < 1 || hook.timeoutMs > 10_000)
    )
      throw new Error('Hook timeout must be 1–10000 ms');
    if (hook.priority !== undefined && !Number.isFinite(hook.priority))
      throw new Error('Invalid hook priority');
    // Type erasure stays internal; registration checks preserve the payload/result relation.
    const registered = { ...hook } as RuntimeHook;
    this.hooks.push(registered);
    return () => {
      const index = this.hooks.indexOf(registered);
      if (index >= 0) this.hooks.splice(index, 1);
    };
  }

  list(): Array<{ id: string; on: RuntimeHookPhase; mode: 'guard' | 'observe' }> {
    return this.hooks.map(({ id, on, mode }) => ({ id, on, mode }));
  }
  records(): HookRecord[] {
    return this.history.map((r) => ({ ...r }));
  }
  subscribe(listener: (record: HookRecord) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  async run<K extends RuntimeHookPhase>(
    phase: K,
    payload: RuntimeHookPayloads[K],
    context: { taskId: string; surface: Surface; signal?: AbortSignal },
  ): Promise<ResolvedContext[]> {
    context.signal?.throwIfAborted();
    const selected = this.hooks
      .filter((h) => h.on === phase)
      .sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
    const entries: ResolvedContext[] = [];
    const deadline = Date.now() + 10_000;
    let contextBytes = 0;
    for (const hook of selected) {
      context.signal?.throwIfAborted();
      const started = Date.now();
      const key = `${context.taskId}:${phase}:${hook.id}`;
      let outcome: HookRecord['outcome'] = 'continued';
      let rejection: Error | undefined;
      let entered = false;
      try {
        if (this.running.has(key)) throw new Error('Recursive hook invocation');
        const remaining = deadline - Date.now();
        if (remaining <= 0) throw new HookTimeoutError('Hook phase deadline exceeded');
        this.running.add(key);
        entered = true;
        const result = await boundedCall(
          (signal) =>
            hook.handle(snapshot(payload), {
              taskId: context.taskId,
              surface: context.surface,
              signal,
            }),
          Math.min(hook.timeoutMs ?? 750, remaining),
          context.signal,
        );
        if (result === undefined || result.kind === 'continue') {
          /* observer or no-op */
        } else if (result.kind === 'block') {
          if (hook.mode !== 'guard' || typeof result.reason !== 'string' || !result.reason.trim())
            throw new Error('Invalid hook block');
          outcome = 'blocked';
          rejection = new HookBlockedError(hook.id, phase, result.reason.slice(0, 1000));
        } else if (result.kind === 'context') {
          if (phase !== 'message:received' || hook.mode !== 'guard')
            throw new Error('Only message guards can supply context');
          const parsed = ResolvedContextSchema.array().max(16).parse(result.entries);
          contextBytes += new TextEncoder().encode(JSON.stringify(parsed)).byteLength;
          if (entries.length + parsed.length > 16 || contextBytes > MAX_CONTEXT_BYTES)
            throw new Error('Hook context budget exceeded');
          entries.push(...parsed);
          outcome = 'context';
        } else throw new Error('Invalid hook result');
      } catch (error) {
        outcome = context.signal?.aborted
          ? 'cancelled'
          : error instanceof HookTimeoutError
            ? 'timeout'
            : 'error';
        if (context.signal?.aborted) rejection = abortError();
        else if (hook.mode === 'guard')
          rejection = new HookBlockedError(
            hook.id,
            phase,
            `Required hook ${hook.id} ${outcome === 'timeout' ? 'timed out' : 'failed'}. The operation was stopped.`,
          );
      } finally {
        if (entered) this.running.delete(key);
        this.record({
          sequence: ++this.sequence,
          taskId: context.taskId,
          hookId: hook.id,
          phase,
          outcome,
          durationMs: Date.now() - started,
        });
      }
      if (rejection) throw rejection;
    }
    return entries;
  }

  private record(record: HookRecord): void {
    this.history.push(record);
    if (this.history.length > MAX_RECORDS) this.history.shift();
    for (const listener of this.listeners) {
      try {
        listener({ ...record });
      } catch {
        /* diagnostics never change execution */
      }
    }
  }
}

export function snapshot<T>(value: T): DeepReadonly<T> {
  const copy = structuredClone(value);
  const freeze = (item: unknown): void => {
    if (item && typeof item === 'object' && !Object.isFrozen(item)) {
      Object.freeze(item);
      for (const child of Object.values(item)) freeze(child);
    }
  };
  freeze(copy);
  return copy as DeepReadonly<T>;
}
class HookTimeoutError extends Error {}
function abortError(): DOMException {
  return new DOMException('Operation cancelled', 'AbortError');
}

/** Consume late rejection; never accept a result after timeout/abort, even if a handler ignores its signal. */
export async function boundedCall<T>(
  call: (signal: AbortSignal) => T | Promise<T>,
  timeoutMs: number,
  parent?: AbortSignal,
): Promise<T> {
  parent?.throwIfAborted();
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onAbort: (() => void) | undefined;
  try {
    return await Promise.race([
      Promise.resolve().then(() => {
        controller.signal.throwIfAborted();
        return call(controller.signal);
      }),
      new Promise<never>((_, reject) => {
        onAbort = () => {
          controller.abort();
          reject(abortError());
        };
        parent?.addEventListener('abort', onAbort, { once: true });
        timer = setTimeout(() => {
          controller.abort();
          reject(new HookTimeoutError('Hook timed out'));
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (onAbort) parent?.removeEventListener('abort', onAbort);
    controller.abort();
  }
}
