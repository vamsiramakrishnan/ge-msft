import type { CapabilityManifest } from '@ge/contracts';

/**
 * What PowerPoint can read and write. Static today; `PowerPointBridge.getCapabilities`
 * narrows it at runtime against `Office.context.requirements.isSetSupported('PowerPointApi', …)`
 * so unsupported actions are never offered on a thin host.
 */
export const POWERPOINT_CAPABILITIES: CapabilityManifest = {
  surface: 'powerpoint',
  contextKinds: ['selection', 'slide', 'shape', 'document'],
  // Read verbs PowerPoint serves (ADR-0006 closure): `outline` via `captureDocState` (slide
  // inventory), addressable `read <slide:N>` via `readRange`, `search` via `searchDocument`
  // (slide-text scan). Each backed by a real bridge port; the universal context-attach port
  // (`listContext`/`resolveContext`) is separate and not a read verb.
  reads: ['outline', 'read', 'search'],
  actuations: [
    {
      kind: 'insert-slide',
      surface: 'powerpoint',
      title: 'Add slide',
      description: 'Compose a grounded slide into the deck (title + bullets).',
      reversible: true,
      appliesTo: ['slide'],
    },
    // NOTE: `set-speaker-notes` was advertised but its `actuate()` case ALWAYS degrades — this
    // Office.js typings version has no `Slide.notes`/notesSlide write path. Un-advertised (ADR-0006
    // phantom: advertised-but-never-actuates). Re-add once the host typings expose a notes writer.
  ],
};
