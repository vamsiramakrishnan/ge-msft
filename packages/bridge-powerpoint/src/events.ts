import { type EventOrigin, type HostEvent } from '@ge/triggers';

/**
 * Pure mappers from raw PowerPoint event primitives into our `HostEvent` union — no Office.js
 * here, so it's unit-testable. PowerPoint in this Office.js typings has no object-model
 * selection/change event; the `PowerPointBridge.watch` host wiring registers the Office-level
 * `Office.EventType.DocumentSelectionChanged` (shape/text selection) and `ActiveViewChanged`
 * (slide navigation) and hands the extracted primitives to these builders. Surface is
 * hard-coded `'powerpoint'`.
 *
 * Neither Office-level event carries a coauthoring source, so `origin` is always `'local'`.
 */

/** Build a `selection-changed` HostEvent. No coauthor source on PowerPoint → always local. */
export function selectionChanged(
  preview?: string,
): Extract<HostEvent, { type: 'selection-changed' }> {
  const event: Extract<HostEvent, { type: 'selection-changed' }> = {
    type: 'selection-changed',
    surface: 'powerpoint',
    origin: 'local',
  };
  if (preview !== undefined && preview.length > 0) event.preview = preview;
  return event;
}

/**
 * Build a `document-changed` HostEvent. PowerPoint surfaces slide navigation as
 * `ActiveViewChanged`; we map it to a local document change (the active slide moved). Origin is
 * a parameter for symmetry with the other bridges, defaulting to `'local'`.
 */
export function documentChanged(
  origin: EventOrigin = 'local',
): Extract<HostEvent, { type: 'document-changed' }> {
  return { type: 'document-changed', surface: 'powerpoint', origin };
}
