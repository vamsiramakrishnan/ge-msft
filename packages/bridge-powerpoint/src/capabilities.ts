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
    {
      kind: 'set-shape-text',
      surface: 'powerpoint',
      title: 'Replace shape text',
      description: 'Replace text in an explicitly addressed slide shape or text box.',
      reversible: true,
      appliesTo: ['shape'],
    },
    {
      kind: 'add-shape',
      surface: 'powerpoint',
      title: 'Add shape',
      description: 'Add a text box, geometric shape, or line to a slide with explicit geometry.',
      reversible: true,
      appliesTo: ['slide'],
    },
    {
      kind: 'format-shape',
      surface: 'powerpoint',
      title: 'Format shape',
      description: 'Change fill, line, font, or z-order of an explicitly addressed shape in place.',
      reversible: true,
      appliesTo: ['shape'],
    },
    {
      kind: 'add-table-slide',
      surface: 'powerpoint',
      title: 'Add slide table',
      description: 'Insert a small native table seeded from a value grid onto an addressed slide.',
      reversible: true,
      appliesTo: ['slide', 'table'],
    },
    // NOTE: `set-speaker-notes` was advertised but its `actuate()` case ALWAYS degrades — this
    // Office.js typings version has no `Slide.notes`/notesSlide write path. Un-advertised (ADR-0006
    // phantom: advertised-but-never-actuates). Re-add once the host typings expose a notes writer.
    // The same rule keeps `insert-image` un-advertised: this typings version exposes NO PowerPoint
    // image-insertion write path (only Excel's ShapeCollection.addImage; PPT's getImageAsBase64 is
    // read-only), so there is no typed host call for the bridge to make.
    // NOTE: `set-speaker-notes` was advertised but its `actuate()` case ALWAYS degrades — this
    // Office.js typings version has no `Slide.notes`/notesSlide write path. Un-advertised (ADR-0006
    // phantom: advertised-but-never-actuates). Re-add once the host typings expose a notes writer.
  ],
};
