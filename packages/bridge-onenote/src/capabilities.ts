import type { CapabilityManifest } from '@ge/contracts';

/**
 * What OneNote can read and write. OneNote is web-only and ships with a narrow API surface
 * (`OneNoteApi`), so the capability map is intentionally small. `OneNoteBridge.getCapabilities`
 * narrows it at runtime against `Office.context.requirements.isSetSupported('OneNoteApi', …)`.
 */
export const ONENOTE_CAPABILITIES: CapabilityManifest = {
  surface: 'onenote',
  contextKinds: ['page', 'document'],
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
