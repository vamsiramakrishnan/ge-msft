import type { CapabilityManifest } from '@ge/contracts';

/**
 * What OneNote can read and write. OneNote is web-only and ships with a narrow API surface
 * (`OneNoteApi`), so the capability map is intentionally small. `OneNoteBridge.getCapabilities`
 * narrows it at runtime against `Office.context.requirements.isSetSupported('OneNoteApi', …)`.
 */
export const ONENOTE_CAPABILITIES: CapabilityManifest = {
  surface: 'onenote',
  contextKinds: ['page', 'document'],
  // Read verbs OneNote serves (ADR-0006 closure): `outline` via `captureDocState` (active-page
  // title + paragraph outline), whole-page `read` (the runtime's empty-selector read → the same
  // `captureDocState`, since a OneNote page has no addressable sub-range), and `search` via
  // `searchDocument` (page-paragraph scan). OneNote is web-only with a narrow API, so the
  // "document" is the active page. Context attach (`listContext`/`resolveContext`) is separate.
  reads: ['outline', 'read', 'search'],
  actuations: [
    {
      kind: 'append-page',
      surface: 'onenote',
      title: 'Synthesize onto a page',
      description: 'Add a new page of grounded synthesis, one inline citation tag per claim.',
      reversible: true,
      appliesTo: ['page'],
    },
  ],
};
