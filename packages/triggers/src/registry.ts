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
  handle: (event: HostEvent) => TriggerOutcome | Promise<TriggerOutcome>;
}

function matches(trigger: Trigger, event: HostEvent): boolean {
  const types = Array.isArray(trigger.on) ? trigger.on : [trigger.on];
  if (!types.includes(event.type)) return false;
  if ((trigger.ignoreRemote ?? true) && eventOrigin(event) === 'remote') return false;
  return trigger.match ? trigger.match(event) : true;
}

/**
 * Holds the trigger set and dispatches events to it. Pure matching + invocation — no timers
 * here (debouncing is an event-source concern; see `debounce`). Two entry points:
 *   • `dispatch` — fire all matching triggers, collect outcomes (ambient/automate).
 *   • `gate`     — run matching triggers as a *gate* (pre-actuation / on-send): the first
 *                  `block` wins, mirroring a PreToolUse hook that can veto an action.
 */
export class TriggerRegistry {
  private readonly triggers: Trigger[] = [];

  register(trigger: Trigger): () => void {
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

  /** Fire every matching trigger; return all outcomes (skips plain 'continue'). */
  async dispatch(event: HostEvent): Promise<TriggerOutcome[]> {
    const outcomes: TriggerOutcome[] = [];
    for (const t of this.triggers) {
      if (!matches(t, event)) continue;
      const outcome = await t.handle(event);
      if (outcome.kind !== 'continue') outcomes.push(outcome);
    }
    return outcomes;
  }

  /**
   * Run matching triggers as a veto gate. The first `block` short-circuits and is returned;
   * otherwise `continue`. Use for `pre-actuation` (before a write lands) and `mail-send`.
   */
  async gate(event: HostEvent): Promise<TriggerOutcome> {
    for (const t of this.triggers) {
      if (!matches(t, event)) continue;
      const outcome = await t.handle(event);
      if (outcome.kind === 'block') return outcome;
    }
    return CONTINUE;
  }
}
