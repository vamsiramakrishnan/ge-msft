import type { CapabilityManifest } from '@ge/contracts';

/**
 * What PowerPoint can read and write. Static today; `PowerPointBridge.getCapabilities`
 * narrows it at runtime against `Office.context.requirements.isSetSupported('PowerPointApi', …)`
 * so unsupported actions are never offered on a thin host.
 */
export const POWERPOINT_CAPABILITIES: CapabilityManifest = {
  surface: 'powerpoint',
  contextKinds: ['selection', 'slide', 'shape', 'document'],
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
      kind: 'set-speaker-notes',
      surface: 'powerpoint',
      title: 'Set speaker notes',
      description: "Replace the slide's speaker notes with grounded talking points.",
      reversible: true,
      appliesTo: ['slide'],
    },
  ],
};
