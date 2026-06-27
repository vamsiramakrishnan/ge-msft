import { describe, it, expect } from 'vitest';
import { IntentSchema } from './intent.js';
import type { CapabilityManifest } from './capability.js';
import { INTENT_REQUIRES, intentsForManifest } from './intent-capability.js';

function manifest(kinds: CapabilityManifest['actuations'][number]['kind'][]): CapabilityManifest {
  return {
    surface: 'word',
    contextKinds: [],
    actuations: kinds.map((kind) => ({ kind, surface: 'word', title: kind, reversible: true })),
  };
}

describe('INTENT_REQUIRES', () => {
  it('maps every Intent (no missing entry)', () => {
    expect(Object.keys(INTENT_REQUIRES).sort()).toEqual([...IntentSchema.options].sort());
  });

  it('the chat verbs require no actuation; the specialist verbs do', () => {
    expect(INTENT_REQUIRES.ask).toEqual([]);
    expect(INTENT_REQUIRES.summarize).toEqual([]);
    expect(INTENT_REQUIRES.explain).toEqual([]);
    expect(INTENT_REQUIRES.rewrite.length).toBeGreaterThan(0);
    expect(INTENT_REQUIRES.review.length).toBeGreaterThan(0);
    expect(INTENT_REQUIRES.visualize).toEqual(['insert-chart']);
    expect(INTENT_REQUIRES.draft.length).toBeGreaterThan(0);
    expect(INTENT_REQUIRES.notes.length).toBeGreaterThan(0);
  });
});

describe('intentsForManifest', () => {
  it('always allows the chat verbs, even with zero actuations', () => {
    expect(intentsForManifest(manifest([]))).toEqual(['ask', 'summarize', 'explain']);
  });

  it('allows rewrite/review when their in-document writes are advertised', () => {
    const allowed = intentsForManifest(manifest(['tracked-change', 'add-comment']));
    expect(allowed).toEqual(
      expect.arrayContaining(['ask', 'summarize', 'explain', 'rewrite', 'review']),
    );
    // No slide/page/mail/post capability → draft/notes are not offered.
    expect(allowed).not.toContain('draft');
    expect(allowed).not.toContain('notes');
  });

  it('satisfies review by EITHER a comment or a tracked change', () => {
    expect(intentsForManifest(manifest(['add-comment']))).toContain('review');
    expect(intentsForManifest(manifest(['tracked-change']))).toContain('review');
  });

  it('satisfies rewrite by any of its reversible writes', () => {
    expect(intentsForManifest(manifest(['write-cells']))).toContain('rewrite');
    expect(intentsForManifest(manifest(['fill-content-control']))).toContain('rewrite');
    expect(intentsForManifest(manifest(['replace-selection']))).toContain('rewrite');
    // A comment-only manifest cannot land a rewrite.
    expect(intentsForManifest(manifest(['add-comment']))).not.toContain('rewrite');
  });

  it('allows visualize only when chart insertion is advertised', () => {
    expect(intentsForManifest(manifest(['insert-chart']))).toContain('visualize');
    expect(intentsForManifest(manifest(['write-cells', 'create-table']))).not.toContain(
      'visualize',
    );
  });

  it('maps surface-specific generators (draft/notes) to their actuations', () => {
    expect(intentsForManifest(manifest(['insert-slide']))).toContain('draft');
    expect(intentsForManifest(manifest(['append-page']))).toContain('draft');
    expect(intentsForManifest(manifest(['create-mail']))).toContain('draft');
    expect(intentsForManifest(manifest(['reply-mail']))).toContain('draft');
    expect(intentsForManifest(manifest(['post-message']))).toContain('notes');
    expect(intentsForManifest(manifest(['post-card']))).toContain('notes');
  });
});
