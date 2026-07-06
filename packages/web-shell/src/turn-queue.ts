import type { ResolvedGrounding } from '@ge/gemini-client';
import { assertNever } from './assert-never.js';

/**
 * A TURN queued while another is in flight (Finding #3). The single-slot queue is LATEST-WINS, but
 * it is TYPED by mode so a queued turn drains through the SAME route it was requested on — a queued
 * `commands` turn re-enters the fail-closed plan/approval loop, NEVER plain grounded chat. Collapsing
 * every queued turn into a string and re-sending it through `send()` (the prior `pendingQuery?:
 * string`) silently downgraded a queued write/annotation turn into chat; this discriminated union
 * makes that impossible — the mode is carried through the drain.
 */
export type QueuedTurn =
  | { mode: 'ask'; query: string; grounding?: ResolvedGrounding }
  | { mode: 'commands'; task: string; grounding?: ResolvedGrounding }
  | { mode: 'direct-commands'; program: string }
  | { mode: 'skill'; name: string; args: Record<string, string> };

/** The mode-preserving drain handlers — one per {@link QueuedTurn} variant. */
export interface DrainHandlers {
  ask(query: string, grounding?: ResolvedGrounding): void;
  commands(task: string, grounding?: ResolvedGrounding): void;
  directCommands(program: string): void;
  skill(name: string, args: Record<string, string>): void;
}

/**
 * The single-slot, latest-wins turn queue. Extracted from `PanelController` (E-full) as the cohesive
 * "what runs next" responsibility: it holds at most one queued turn and drains it back through the
 * route it was requested on — so a queued mode is preserved end-to-end and can never be downgraded.
 */
export class TurnQueue {
  private pending: QueuedTurn | undefined;

  /** Hold a turn to run when the current one settles (replaces any prior queued turn). */
  enqueue(turn: QueuedTurn): void {
    this.pending = turn;
  }

  /** Whether a turn is currently queued (for assertions/introspection). */
  get queued(): boolean {
    return this.pending !== undefined;
  }

  /**
   * Drain the queued turn (if any) through its matching handler, clearing the slot FIRST so the
   * dispatch can't re-enqueue itself synchronously. Exhaustive via {@link assertNever}.
   */
  drain(handlers: DrainHandlers): void {
    const next = this.pending;
    if (!next) return;
    this.pending = undefined;
    switch (next.mode) {
      case 'ask':
        handlers.ask(next.query, next.grounding);
        return;
      case 'commands':
        handlers.commands(next.task, next.grounding);
        return;
      case 'direct-commands':
        handlers.directCommands(next.program);
        return;
      case 'skill':
        handlers.skill(next.name, next.args);
        return;
      default:
        assertNever(next);
    }
  }
}
