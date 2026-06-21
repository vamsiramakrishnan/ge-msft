import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Outlook can read and write. Static today; `OutlookBridge.getCapabilities` could narrow
 * it at runtime against `Office.context.requirements.isSetSupported('Mailbox', …)` so actions
 * unavailable on a thin platform (e.g. compose-only hosts) are never offered.
 */
export const OUTLOOK_CAPABILITIES: CapabilityManifest = {
  surface: 'outlook',
  contextKinds: ['mail-item', 'mail-thread', 'attachment'],
  actuations: [
    {
      kind: 'reply-mail',
      surface: 'outlook',
      title: 'Draft grounded reply',
      description: 'Open a reply form pre-filled with a grounded, reviewable draft.',
      reversible: true,
      appliesTo: ['mail-item'],
    },
    {
      kind: 'create-mail',
      surface: 'outlook',
      title: 'Draft new message',
      description: 'Open a new draft message pre-filled with grounded content.',
      reversible: true,
    },
  ],
};
