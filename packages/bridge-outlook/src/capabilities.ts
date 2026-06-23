import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Outlook can read and write. Static today; `OutlookBridge.getCapabilities` could narrow
 * it at runtime against `Office.context.requirements.isSetSupported('Mailbox', …)` so actions
 * unavailable on a thin platform (e.g. compose-only hosts) are never offered.
 */
export const OUTLOOK_CAPABILITIES: CapabilityManifest = {
  surface: 'outlook',
  contextKinds: ['mail-item', 'mail-thread', 'attachment'],
  // Read verbs Outlook serves (ADR-0006 closure): whole-item `read` via `captureDocState` (a mail
  // item has no addressable sub-range, so the "document" is the single active item — subject + from
  // + leading body lines), and `search` via `searchDocument` (a body-line scan). No `outline`: a
  // mail item has no heading structure to outline. Context attach (`listContext`/`resolveContext`)
  // is separate and not a read verb.
  reads: ['read', 'search'],
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
