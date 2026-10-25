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
  // `substitutes` on the draft edits: what can stand in when the pane is on a received message. A
  // reply form takes a body (and a file only through `mail … attach=`, never as a stand-in for a
  // refused add-attachment); a new message built from a command takes a subject and a body (never
  // recipients: it is left for the user to address).
  actuations: [
    {
      kind: 'reply-mail',
      surface: 'outlook',
      title: 'Draft grounded reply (can carry a Word or text file)',
      description:
        'Open a reply form pre-filled with a grounded, reviewable draft, with a Word or text file attached when asked.',
      reversible: true,
      appliesTo: ['mail-item'],
    },
    {
      kind: 'create-mail',
      surface: 'outlook',
      title: 'Draft new email',
      description:
        'Open a new message form pre-filled with a grounded subject + body (unaddressed).',
      reversible: true,
      appliesTo: ['mail-item'],
    },
    {
      kind: 'add-attachment',
      surface: 'outlook',
      title: 'Attach file to draft',
      description:
        'Attach a file to the open draft — name="X.docx" text="…" builds a Word file, .txt a text file — or an https link or mail item; nothing sends.',
      reversible: true,
      appliesTo: ['mail-item', 'attachment'],
      substitutes: [],
    },
    {
      kind: 'set-body',
      surface: 'outlook',
      title: 'Replace draft body',
      description: 'Replace the body of the open draft as HTML or text; nothing sends.',
      reversible: true,
      appliesTo: ['mail-item'],
      substitutes: ['reply-mail', 'create-mail'],
    },
    {
      kind: 'set-recipients',
      surface: 'outlook',
      title: 'Set draft recipients',
      description: 'Set (replace) or add To/Cc/Bcc on the open draft; send stays with the user.',
      reversible: true,
      appliesTo: ['mail-item', 'person'],
      substitutes: [],
    },
    {
      kind: 'set-subject',
      surface: 'outlook',
      title: 'Set draft subject',
      description: 'Replace the subject of the open draft; nothing sends.',
      reversible: true,
      appliesTo: ['mail-item'],
      substitutes: ['create-mail'],
    },
  ],
};
