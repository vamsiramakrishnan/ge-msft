import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Outlook can read and write. Static today; `OutlookBridge.getCapabilities` could narrow
 * it at runtime against `Office.context.requirements.isSetSupported('Mailbox', …)` so actions
 * unavailable on a thin platform (e.g. compose-only hosts) are never offered.
 */
export const OUTLOOK_CAPABILITIES: CapabilityManifest = {
  surface: 'outlook',
  contextKinds: ['mail-item', 'mail-thread', 'attachment'],
  // No addressable read verbs: Outlook captures the active mail item via the universal
  // `listContext`/`resolveContext` port, but has no `outline`/`read`/`search` read port today
  // (no `captureDocState`/`readRange`/`searchDocument`). ADR-0006 closure: advertise nothing.
  reads: [],
  actuations: [
    {
      kind: 'reply-mail',
      surface: 'outlook',
      title: 'Draft grounded reply',
      description: 'Open a reply form pre-filled with a grounded, reviewable draft.',
      reversible: true,
      appliesTo: ['mail-item'],
    },
    // NOTE: `create-mail` was advertised but `actuate()` handles only `reply-mail` (ADR-0006
    // phantom). Un-advertised rather than implemented. Re-add with a `displayNewMessageForm`
    // `actuate()` case + a CLI verb when a "draft new message" flow is actually built.
  ],
};
