import { coauthorOrigin, type EventOrigin, type HostEvent } from '@ge/triggers';

/**
 * Pure mappers from raw Excel event primitives into our `HostEvent` union — no Office.js
 * here, so it's unit-testable. The `ExcelBridge.watch` host wiring extracts the primitives
 * (`address`, `source`, `commentId`) from the Office.js event args inside `Excel.run` and
 * hands them to these builders. Surface is hard-coded `'excel'`.
 *
 * Treat all extracted strings (addresses, comment ids) as untrusted host data: these
 * builders only carry them, they never interpret them as instructions.
 */

/** The shape Excel's coauthoring `source` field takes (`Excel.EventSource` is `"Local"|"Remote"`). */
export type ExcelEventSourceLike = string | undefined;

/**
 * Map Excel's coauthoring source (`Excel.EventSource` string enum `"Local"`/`"Remote"`) to our
 * `EventOrigin`. Thin surface-named wrapper over the shared `coauthorOrigin` rule.
 */
export function deriveOrigin(source: ExcelEventSourceLike): EventOrigin {
  return coauthorOrigin(source);
}

/** Build a `selection-changed` HostEvent. Selection has no coauthor source → always local. */
export function selectionChanged(
  address?: string,
): Extract<HostEvent, { type: 'selection-changed' }> {
  const event: Extract<HostEvent, { type: 'selection-changed' }> = {
    type: 'selection-changed',
    surface: 'excel',
    origin: 'local',
  };
  if (address !== undefined && address.length > 0) event.preview = address;
  return event;
}

/** Build a `document-changed` HostEvent from an already-derived origin. */
export function documentChanged(
  origin: EventOrigin,
): Extract<HostEvent, { type: 'document-changed' }> {
  return { type: 'document-changed', surface: 'excel', origin };
}

/** Build a `comment-added` HostEvent for a single comment id. */
export function commentAdded(
  commentId: string,
  origin: EventOrigin,
  text?: string,
): Extract<HostEvent, { type: 'comment-added' }> {
  const event: Extract<HostEvent, { type: 'comment-added' }> = {
    type: 'comment-added',
    surface: 'excel',
    origin,
    commentId,
  };
  if (text !== undefined) event.text = text;
  return event;
}
