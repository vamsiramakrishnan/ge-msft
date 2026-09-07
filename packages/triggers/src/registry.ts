import {
  CONTINUE,
  eventOrigin,
  type EventType,
  type HostEvent,
  type TriggerOutcome,
} from './event.js';

/**
 * A registered reaction. `on` selects event types; `match` further filters; `ignoreRemote`
 * (default true) drops coauthor/own-write events so the system never reacts to its own edits.
 * `handle` returns an outcome. Triggers are pure functions of an event — easy to test, compose,
 * and ship as plugins.
 */
export interface Trigger {
  id: string;
  on: EventType | EventType[];
  match?: (event: HostEvent) => boolean;
  ignoreRemote?: boolean;
  priority?: number;
  timeoutMs?: number;
  handle: (
    event: HostEvent,
    context: { signal: AbortSignal },
  ) => TriggerOutcome | Promise<TriggerOutcome>;
}

function matches(trigger: Trigger, event: HostEvent): boolean {
  const types = Array.isArray(trigger.on) ? trigger.on : [trigger.on];
  if (!types.includes(event.type)) return false;
  if ((trigger.ignoreRemote ?? true) && eventOrigin(event) === 'remote') return false;
  return trigger.match ? trigger.match(event) : true;
}

/**
 * Holds the trigger set and dispatches events with bounded matching and invocation.
 * Debouncing is an event-source concern; see `debounce`. Two entry points:
 *   • `dispatch` — fire all matching triggers, collect outcomes (ambient/automate).
 *   • `gate`     — run matching triggers as a *gate* (pre-actuation / on-send): the first
 *                  `block` wins, mirroring a PreToolUse hook that can veto an action.
 */
export class TriggerRegistry {
  private readonly triggers: Trigger[] = [];
  private readonly listeners = new Set<(record: TriggerDiagnostic) => void>();

  subscribe(listener: (record: TriggerDiagnostic) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
  private report(trigger: Trigger, event: HostEvent): void {
    for (const listener of this.listeners) {
      try {
        listener({ triggerId: trigger.id, event: event.type, outcome: 'failed' });
      } catch {
        /* observational */
      }
    }
  }

  register(trigger: Trigger): () => void {
    if (
      !/^[a-zA-Z0-9][a-zA-Z0-9._:/-]{0,119}$/.test(trigger.id) ||
      typeof trigger.handle !== 'function'
    )
      throw new Error('Invalid trigger definition');
    if (this.triggers.some((t) => t.id === trigger.id))
      throw new Error(`Duplicate trigger id: ${trigger.id}`);
    if (this.triggers.length >= 64) throw new Error('Trigger limit reached');
    if (
      trigger.timeoutMs !== undefined &&
      (!Number.isFinite(trigger.timeoutMs) || trigger.timeoutMs < 1 || trigger.timeoutMs > 10_000)
    )
      throw new Error('Invalid trigger deadline');
    if (trigger.priority !== undefined && !Number.isFinite(trigger.priority))
      throw new Error('Invalid trigger priority');
    trigger = { ...trigger, on: Array.isArray(trigger.on) ? [...trigger.on] : trigger.on };
    this.triggers.push(trigger);
    return () => {
      const i = this.triggers.indexOf(trigger);
      if (i >= 0) this.triggers.splice(i, 1);
    };
  }

  registerAll(triggers: Trigger[]): void {
    for (const t of triggers) this.register(t);
  }

  get size(): number {
    return this.triggers.length;
  }

  private ordered(): Trigger[] {
    return [...this.triggers].sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
  }

  /** Fire every matching trigger; return all outcomes (skips plain 'continue'). */
  async dispatch(event: HostEvent): Promise<TriggerOutcome[]> {
    const outcomes: TriggerOutcome[] = [];
    const deadline = Date.now() + 5000;
    for (const t of this.ordered()) {
      try {
        const outcome = await invoke(t, event, deadline);
        if (outcome.kind !== 'continue') outcomes.push(outcome);
      } catch {
        this.report(t, event);
      }
    }
    return outcomes;
  }

  /**
   * Run matching triggers as a veto gate. The first `block` short-circuits and is returned;
   * otherwise `continue`. Use for `pre-actuation` (before a write lands) and `mail-send`.
   */
  async gate(event: HostEvent): Promise<TriggerOutcome> {
    const deadline = Date.now() + 5000;
    for (const t of this.ordered()) {
      try {
        const outcome = await invoke(t, event, deadline);
        if (outcome.kind === 'block') return outcome;
      } catch {
        this.report(t, event);
        return {
          kind: 'block',
          reason: `Required check ${t.id} could not complete. Try again before applying this action.`,
        };
      }
    }
    return CONTINUE;
  }
}

/** Include matching in the deadline and isolate each handler from the original request and other handlers. */
export interface TriggerDiagnostic {
  triggerId: string;
  event: HostEvent['type'];
  outcome: 'failed';
}

async function invoke(
  trigger: Trigger,
  event: HostEvent,
  deadline: number,
): Promise<TriggerOutcome> {
  const types = Array.isArray(trigger.on) ? trigger.on : [trigger.on];
  if (!types.includes(event.type)) return CONTINUE;
  if ((trigger.ignoreRemote ?? true) && eventOrigin(event) === 'remote') return CONTINUE;
  const remaining = deadline - Date.now();
  if (remaining <= 0) throw new Error('Trigger dispatch deadline exceeded');
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      Promise.resolve().then(async () => {
        const copy = structuredClone(event);
        const freeze = (value: unknown): void => {
          if (value && typeof value === 'object' && !Object.isFrozen(value)) {
            Object.freeze(value);
            for (const child of Object.values(value)) freeze(child);
          }
        };
        freeze(copy);
        if (!matches(trigger, copy)) return CONTINUE;
        const result = await trigger.handle(copy, { signal: controller.signal });
        if (!result || !['continue', 'block', 'suggest', 'automate'].includes(result.kind))
          throw new Error('Invalid trigger result');
        if (result.kind === 'block' && (typeof result.reason !== 'string' || !result.reason.trim()))
          throw new Error('Invalid block reason');
        if (result.kind === 'suggest' && (typeof result.title !== 'string' || !result.title.trim()))
          throw new Error('Invalid suggestion');
        if (
          result.kind === 'automate' &&
          (typeof result.query !== 'string' || !result.query.trim())
        )
          throw new Error('Invalid automation');
        return structuredClone(result);
      }),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () => {
            controller.abort();
            reject(new Error('Trigger timed out'));
          },
          Math.min(trigger.timeoutMs ?? 750, remaining),
        );
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    controller.abort();
  }
}
