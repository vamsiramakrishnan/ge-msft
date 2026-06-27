import type { ContextRef } from '@ge/contracts';
import { coauthorOrigin, type EventOrigin, type HostEvent } from '@ge/triggers';

/**
 * Pure mapping from raw Office.js event payloads into `HostEvent`s — no Office.js calls here,
 * so it's unit-testable. The `WordBridge` registers the host event handlers (`Word.run`,
 * `Office.context`) and hands the already-extracted primitives to these builders, keeping the
 * wiring thin. Surface is hard-coded to `'word'`.
 *
 * Coauthoring: Word content events carry a `source` of `Word.EventSource` ('Local' | 'Remote').
 * We derive `'remote'` only when the host explicitly says remote; everything else (including an
 * unknown/missing source) defaults to `'local'`, so the trigger registry — which drops `remote`
 * by default — never wrongly suppresses a genuine local edit.
 */

/**
 * Map Word's coauthoring `EventSource` (`'Local'`/`'Remote'`) to our `EventOrigin`. Thin
 * surface-named wrapper over the shared `coauthorOrigin` rule; takes `unknown` because handler
 * args are untrusted at the boundary.
 */
export function originFromWordSource(source: unknown): EventOrigin {
  return coauthorOrigin(source);
}

/** Build a `selection-changed` HostEvent. Selection has no coauthor source → always `'local'`. */
export function selectionChangedEvent(preview?: string, ref?: ContextRef): HostEvent {
  return {
    type: 'selection-changed',
    surface: 'word',
    origin: 'local',
    ...(preview !== undefined ? { preview } : {}),
    ...(ref !== undefined ? { ref } : {}),
  };
}

/** Build a `document-changed` HostEvent from an already-derived origin. */
export function documentChangedEvent(origin: EventOrigin): HostEvent {
  return { type: 'document-changed', surface: 'word', origin };
}

/** Build a `comment-added` HostEvent. Text is optional (not all hosts surface it). */
export function commentAddedEvent(
  origin: EventOrigin,
  commentId: string,
  text?: string,
): HostEvent {
  return {
    type: 'comment-added',
    surface: 'word',
    origin,
    commentId,
    ...(text !== undefined ? { text } : {}),
  };
}
