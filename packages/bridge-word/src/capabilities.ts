import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Word can read and write. Static today; `WordBridge.getCapabilities` will narrow it
 * at runtime against `Office.context.requirements.isSetSupported('WordApi', …)` so
 * unsupported actions (e.g. annotations on a thin platform) are never offered.
 */
export const WORD_CAPABILITIES: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document', 'paragraph', 'table', 'comment'],
  // Read verbs Word actually serves (ADR-0006 closure): `outline`/`read` (whole-document) via
  // `captureDocState`, `search` via `searchDocument`. No addressable `readRange` port — `read`
  // degrades to the whole-document capture, which is the closure-honest behaviour.
  reads: ['outline', 'read', 'search'],
  actuations: [
    {
      kind: 'insert-text',
      surface: 'word',
      title: 'Insert text',
      description:
        'Insert plain text at the current selection or after an explicit content anchor.',
      reversible: false,
      appliesTo: ['selection', 'paragraph'],
    },
    {
      kind: 'replace-selection',
      surface: 'word',
      title: 'Replace selection',
      description: 'Replace the currently selected text, capturing the prior text as an inverse.',
      reversible: true,
      appliesTo: ['selection'],
    },
    {
      kind: 'insert-ooxml',
      surface: 'word',
      title: 'Insert OOXML',
      description:
        'Insert screened rich OOXML at the current selection or after an explicit content anchor.',
      reversible: false,
      appliesTo: ['selection', 'paragraph'],
    },
    {
      kind: 'tracked-change',
      surface: 'word',
      title: 'Insert as tracked change',
      description: 'Replace the anchored text with a reviewable tracked change.',
      reversible: true,
      appliesTo: ['selection', 'paragraph'],
    },
    {
      kind: 'fill-content-control',
      surface: 'word',
      title: 'Fill content control',
      description:
        'Replace a known content control value, capturing the prior value as an inverse.',
      reversible: true,
      appliesTo: ['document'],
    },
    {
      kind: 'add-comment',
      surface: 'word',
      title: 'Add comment',
      description: 'Attach a new comment anchored on the matched text.',
      reversible: true,
      appliesTo: ['selection', 'paragraph'],
    },
    {
      kind: 'comment-reply',
      surface: 'word',
      title: 'Reply & resolve comment',
      reversible: true,
      appliesTo: ['comment'],
    },
  ],
};
