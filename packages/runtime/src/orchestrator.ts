import {
  debounce,
  type HostEvent,
  type Scheduler,
  type TriggerRegistry,
  type Debounced,
} from '@ge/triggers';
import type { DocBridge } from './bridge.js';
import { boundedCall } from './hooks.js';

/**
 * Wires a surface's host events into the trigger engine and routes the outcomes. This is the
 * "react to events" spine: bridge.watch → (debounce high-frequency events) → registry.dispatch
 * → outcomes. `suggest` becomes an ambient card; `automate` hands a grounded query back to the
 * app to run through an AssistSession. Gates (`block`) are handled at the actuation boundary, not
 * here. Start/stop is idempotent and unsubscribes cleanly.
 */
export interface OrchestratorHandlers {
  /**
   * Every host event (after debouncing) — the context path. Wire this to
   * `AssistSession.ingest` so events *construct* the working-context brief without running the
   * assistant. Fires for all events, independent of whether any trigger matches.
   */
  onContext?(event: HostEvent, context: { signal: AbortSignal }): void | Promise<void>;
  /** An ambient suggestion to render (non-intrusive). */
  onSuggest?(outcome: { title: string; detail?: string; query?: string }): void;
  /** A grounded query the app should run (e.g. via AssistSession.ask). */
  onAutomate?(query: string): void;
  onError?(code: string): void;
}

export interface OrchestratorOptions {
  /** Debounce window (ms) for high-frequency content events. Default 400. */
  debounceMs?: number;
  scheduler?: Scheduler;
  emitLifecycle?: boolean;
}

const HIGH_FREQUENCY: ReadonlySet<HostEvent['type']> = new Set([
  'selection-changed',
  'document-changed',
]);

export class Orchestrator {
  private unsubscribe: (() => void) | undefined;
  private readonly debouncers = new Map<string, Debounced<[HostEvent]>>();
  private epoch = 0;
  private active = false;
  private pending = 0;
  private queue: Promise<void> = Promise.resolve();
  private enqueueEvent?: (event: HostEvent) => boolean;
  private lifetime?: AbortController;
  private report(code: string): void {
    try {
      this.handlers.onError?.(code);
    } catch {
      /* host callbacks cannot break dispatch */
    }
  }

  constructor(
    private readonly bridge: DocBridge,
    private readonly registry: TriggerRegistry,
    private readonly handlers: OrchestratorHandlers = {},
    private readonly options: OrchestratorOptions = {},
  ) {}

  /** Begin watching host events. No-op if the bridge can't observe events or already started. */
  start(): void {
    if (this.active) return;
    this.active = true;
    const epoch = ++this.epoch;
    this.lifetime = new AbortController();
    const enqueue = (event: HostEvent): boolean => {
      if (!this.active || epoch !== this.epoch) return false;
      if (this.pending >= 64) {
        this.report('host_event_backlog');
        return false;
      }
      const admitted = structuredClone(event);
      this.pending++;
      this.queue = this.queue
        .then(async () => {
          if (this.active && epoch === this.epoch) await this.route(admitted, epoch);
        })
        .catch(() => {
          if (this.active && epoch === this.epoch) this.report('host_event_failed');
        })
        .finally(() => {
          this.pending--;
        });
      return true;
    };
    this.enqueueEvent = enqueue;
    if (this.options.emitLifecycle)
      enqueue({ type: 'session-start', surface: this.bridge.surface });
    try {
      this.unsubscribe = this.bridge.watch?.((event) => {
        if (!this.active || epoch !== this.epoch) return;
        if (!HIGH_FREQUENCY.has(event.type)) {
          enqueue(event);
          return;
        }
        let handler = this.debouncers.get(event.type);
        if (!handler) {
          handler = debounce(enqueue, this.options.debounceMs ?? 400, this.options.scheduler);
          this.debouncers.set(event.type, handler);
        }
        handler(event);
      });
    } catch {
      this.report('host_watch_failed');
    }
  }

  stop(): void {
    this.active = false;
    this.lifetime?.abort();
    this.enqueueEvent = undefined;
    this.epoch++;
    for (const handler of this.debouncers.values()) handler.cancel();
    this.debouncers.clear();
    try {
      this.unsubscribe?.();
    } catch {
      this.report('host_unsubscribe_failed');
    } finally {
      this.unsubscribe = undefined;
    }
  }

  /** Tests and explicit callers may drain admitted events without sleeping. */
  async idle(): Promise<void> {
    await this.queue;
  }

  /** Ingress for trusted event-source adapters (e.g. authenticated Graph change feeds). */
  publish(event: HostEvent): boolean {
    if (!this.active || !this.enqueueEvent) return false;
    return this.enqueueEvent(structuredClone(event));
  }

  private async route(event: HostEvent, epoch: number): Promise<void> {
    // Context path first: every event constructs the working brief (cheap, no model call).
    await boundedCall(
      (signal) => this.handlers.onContext?.(structuredClone(event), { signal }),
      5000,
      this.lifetime?.signal,
    );
    if (!this.active || epoch !== this.epoch) return;
    for (const outcome of await this.registry.dispatch(event)) {
      if (!this.active || epoch !== this.epoch) return;
      if (outcome.kind === 'suggest') {
        this.handlers.onSuggest?.({
          title: outcome.title,
          ...(outcome.detail ? { detail: outcome.detail } : {}),
          ...(outcome.query ? { query: outcome.query } : {}),
        });
      } else if (outcome.kind === 'automate') {
        this.handlers.onAutomate?.(outcome.query);
      }
    }
  }
}
