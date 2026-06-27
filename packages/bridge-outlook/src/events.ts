import type { HostEvent, TriggerOutcome } from '@ge/triggers';

/**
 * Pure HostEvent builders + the on-send gate decision — NO Office.js here, so this is the
 * unit-testable heart of the Outlook event source. The bridge (`outlook-bridge.ts`) and the
 * on-send glue (`on-send.ts`) call these after reading the host item; this module never reads
 * the mailbox itself. The mail's subject/body are untrusted data and are NOT inspected here;
 * the only field that crosses into an event is the opaque `itemId`.
 */

/**
 * The user is viewing/editing a draft (compose mode, or the task pane switched to a draft).
 * `id` is the saved item id when one exists — drafts are often unsaved, so it's optional.
 */
export function composeEvent(id?: string): Extract<HostEvent, { type: 'mail-compose' }> {
  return id ? { type: 'mail-compose', id } : { type: 'mail-compose' };
}

/** A received message became active (read mode). Read items always carry a stable `itemId`. */
export function receivedEvent(id: string): Extract<HostEvent, { type: 'mail-received' }> {
  return { type: 'mail-received', id };
}

/** The user pressed Send — the gate candidate fed to `TriggerRegistry.gate`. */
export function sendEvent(id?: string): Extract<HostEvent, { type: 'mail-send' }> {
  return id ? { type: 'mail-send', id } : { type: 'mail-send' };
}

/** The pure on-send decision: what the host should do with `event.completed`. */
export interface SendDecision {
  /** Passed straight to Smart Alerts `event.completed({ allowEvent })`. */
  allowEvent: boolean;
  /** The block reason, surfaced to the user via the Smart Alerts dialog. Present only when blocked. */
  message?: string;
}

/**
 * Map a gate `TriggerOutcome` to the on-send decision. A `block` cancels Send and carries its
 * reason; every other outcome (`continue` / `suggest` / `automate`) lets Send proceed — the
 * gate only ever *vetoes*, it never composes on the user's behalf at send time. Pure and
 * strict (no try/catch): the fail-safe wrapper lives in the host factory.
 */
export function decideSend(outcome: TriggerOutcome): SendDecision {
  if (outcome.kind === 'block') {
    return { allowEvent: false, message: outcome.reason };
  }
  return { allowEvent: true };
}
