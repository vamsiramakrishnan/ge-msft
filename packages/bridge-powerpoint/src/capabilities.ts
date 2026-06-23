import type { CapabilityManifest } from '@ge/contracts';

/**
 * What PowerPoint can read and write. Static today; `PowerPointBridge.getCapabilities`
 * narrows it at runtime against `Office.context.requirements.isSetSupported('PowerPointApi', …)`
 * so unsupported actions are never offered on a thin host.
 */
export const POWERPOINT_CAPABILITIES: CapabilityManifest = {
  surface: 'powerpoint',
  contextKinds: ['selection', 'slide', 'shape', 'document'],
  // No addressable read verbs: PowerPoint has no `captureDocState`/`readRange`/`searchDocument`
  // port today, so it advertises no `outline`/`read`/`search` (ADR-0006 closure). Context attach
  // (`listContext`/`resolveContext`) is unaffected — that is the universal port, not a read verb.
  reads: [],
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
