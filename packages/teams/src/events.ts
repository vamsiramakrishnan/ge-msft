import type { HostEvent } from '@ge/triggers';

/**
 * Pure HostEvent builders for the Teams surface — NO TeamsJS here, so this is the unit-testable
 * heart of the Teams event source. The bridge (`teams-bridge.ts`) calls these after reading
 * TeamsJS lifecycle signals; this module never touches the host. Transcript/chat content is
 * untrusted data and is NOT inspected here — the only field that crosses into an event is the
 * opaque meeting `id`.
 */

/**
 * The pane became active for a Teams meeting/chat session. Carries only the surface tag — the
 * runtime primes the session on this.
 */
export function sessionStartEvent(): Extract<HostEvent, { type: 'session-start' }> {
  return { type: 'session-start', surface: 'teams' };
}

/** The pane is tearing down (frame unmounted / meeting frame closed). */
export function sessionEndEvent(): Extract<HostEvent, { type: 'session-end' }> {
  return { type: 'session-end', surface: 'teams' };
}

/**
 * The meeting wrapped (TeamsJS `meeting` end signal). `id` is the opaque meeting/chat id when
 * one is resolvable from the host context; it is the only host value carried into the event.
 */
export function meetingEndedEvent(id: string): Extract<HostEvent, { type: 'meeting-ended' }> {
  return { type: 'meeting-ended', id };
}
