import type {
  ActuationRequest,
  ActuationResult,
  ContextRef,
  EstateSource,
  SourceRef,
  Surface,
} from '@ge/contracts';

/**
 * The event-driven layer (Claude-Code-shaped hooks for Office). Bridges and Graph emit
 * **lifecycle events**; registered **triggers** match them and return an **outcome**.
 * Two planes of events: host/estate activity, and the actuation lifecycle (`pre-actuation`
 * is the `PreToolUse` analog — it can *block* a write before it lands; `post-actuation` is
 * the audit hook).
 *
 * `origin: 'local' | 'remote'` carries the coauthoring signal: triggers ignore `remote`
 * (and the agent's own writes) by default, so the system never reacts to its own edits.
 */
export type EventOrigin = 'local' | 'remote';

export type HostEvent =
  | { type: 'session-start'; surface: Surface }
  | { type: 'session-end'; surface: Surface }
  | {
      type: 'selection-changed';
      surface: Surface;
      origin: EventOrigin;
      preview?: string;
      ref?: ContextRef;
    }
  | { type: 'document-changed'; surface: Surface; origin: EventOrigin }
  | {
      type: 'comment-added';
      surface: Surface;
      origin: EventOrigin;
      commentId: string;
      text?: string;
    }
  | { type: 'mail-received'; id: string }
  | { type: 'mail-compose'; id?: string }
  | { type: 'mail-send'; id?: string } // gate candidate (Outlook on-send)
  | { type: 'meeting-ended'; id: string }
  | { type: 'estate-changed'; source: EstateSource; id: string } // Graph change notification
  | { type: 'pre-actuation'; request: ActuationRequest }
  | { type: 'post-actuation'; request: ActuationRequest; result: ActuationResult };

export type EventType = HostEvent['type'];

/** What a trigger decides. Mirrors hook control flow: continue / block / surface a suggestion / automate. */
export type TriggerOutcome =
  | { kind: 'continue' }
  | { kind: 'block'; reason: string } // gate: refuse the pending action (e.g. on-send)
  | { kind: 'suggest'; title: string; detail?: string; query?: string; sources?: SourceRef[] }
  | { kind: 'automate'; query: string }; // run a grounded assist turn

export const CONTINUE: TriggerOutcome = { kind: 'continue' };

/** Events that carry a coauthoring origin (content events we must not react to when remote). */
export function eventOrigin(e: HostEvent): EventOrigin | undefined {
  return 'origin' in e ? e.origin : undefined;
}

/**
 * Map a host's coauthoring `source` value (Word/Excel `EventSource`, the string `'Local'`/
 * `'Remote'`, or anything) to our `EventOrigin`. Shared across surface bridges so the rule lives
 * in one place: derive `'remote'` ONLY when the host explicitly says remote (case-insensitively);
 * everything else — including a missing/unknown source — is `'local'`, so a genuine local edit is
 * never mis-tagged as a coauthor's (which the registry would then drop).
 */
export function coauthorOrigin(source: unknown): EventOrigin {
  return typeof source === 'string' && source.toLowerCase() === 'remote' ? 'remote' : 'local';
}
