import type { CapabilityManifest } from '@ge/contracts';

/**
 * What Word can read and write. Static today; `WordBridge.getCapabilities` will narrow it
 * at runtime against `Office.context.requirements.isSetSupported('WordApi', …)` so
 * unsupported actions (e.g. annotations on a thin platform) are never offered.
 */
export const WORD_CAPABILITIES: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document', 'paragraph', 'table', 'comment'],
  actuations: [
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
      reversible: true,
    },
    { kind: 'insert-ooxml', surface: 'word', title: 'Insert formatted content', reversible: true },
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
