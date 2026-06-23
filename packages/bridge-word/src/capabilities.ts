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
      kind: 'tracked-change',
      surface: 'word',
      title: 'Insert as tracked change',
      description: 'Replace the anchored text with a reviewable tracked change.',
      reversible: true,
      appliesTo: ['selection', 'paragraph'],
    },
    // NOTE: `insert-ooxml` and `fill-content-control` were advertised here but `actuate()` handled
    // neither (ADR-0006 phantom capabilities). Un-advertised rather than implemented — a surface
    // must never claim what it cannot do. Re-add ONLY alongside an `actuate()` case + a CLI verb.
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
