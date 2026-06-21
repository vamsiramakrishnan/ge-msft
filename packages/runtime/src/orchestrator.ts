import { debounce, type HostEvent, type Scheduler, type TriggerRegistry } from '@ge/triggers';
import type { DocBridge } from './bridge.js';

/**
 * Wires a surface's host events into the trigger engine and routes the outcomes. This is the
 * "react to events" spine: bridge.watch → (debounce high-frequency events) → registry.dispatch
 * → outcomes. `suggest` becomes an ambient card; `automate` hands a grounded query back to the
 * app to run through an AssistSession. Gates (`block`) are handled at the actuation boundary, not
 * here. Start/stop is idempotent and unsubscribes cleanly.
 */
export interface OrchestratorHandlers {
  /** An ambient suggestion to render (non-intrusive). */
  onSuggest?(outcome: { title: string; detail?: string; query?: string }): void;
  /** A grounded query the app should run (e.g. via AssistSession.ask). */
  onAutomate?(query: string): void;
}

export interface OrchestratorOptions {
  /** Debounce window (ms) for high-frequency content events. Default 400. */
  debounceMs?: number;
  scheduler?: Scheduler;
}

const HIGH_FREQUENCY: ReadonlySet<HostEvent['type']> = new Set([
  'selection-changed',
  'document-changed',
]);

export class Orchestrator {
  private unsubscribe: (() => void) | undefined;

  constructor(
    private readonly bridge: DocBridge,
    private readonly registry: TriggerRegistry,
    private readonly handlers: OrchestratorHandlers = {},
    private readonly options: OrchestratorOptions = {},
  ) {}

  /** Begin watching host events. No-op if the bridge can't observe events or already started. */
  start(): void {
    if (this.unsubscribe || !this.bridge.watch) return;
    const debounced = debounce(
      (event: HostEvent) => void this.route(event),
      this.options.debounceMs ?? 400,
      this.options.scheduler,
    );
    this.unsubscribe = this.bridge.watch((event) => {
      if (HIGH_FREQUENCY.has(event.type)) debounced(event);
      else void this.route(event);
    });
  }

  stop(): void {
    this.unsubscribe?.();
    this.unsubscribe = undefined;
  }

  private async route(event: HostEvent): Promise<void> {
    for (const outcome of await this.registry.dispatch(event)) {
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
