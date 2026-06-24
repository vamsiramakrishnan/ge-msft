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
});

describe('intentsForManifest', () => {
  it('always allows assist, even with zero actuations', () => {
    expect(intentsForManifest(manifest([]))).toEqual(['assist']);
  });

  it('allows the in-document edit/review/resolve verbs when their writes are advertised', () => {
    const allowed = intentsForManifest(
      manifest(['tracked-change', 'add-comment', 'comment-reply']),
    );
    expect(allowed).toEqual(
      expect.arrayContaining(['assist', 'review', 'regen-clause', 'resolve-comment']),
    );
    // No slide/page/post capability → those routing intents are not offered.
    expect(allowed).not.toContain('draft-slides');
    expect(allowed).not.toContain('synthesize');
    expect(allowed).not.toContain('meeting-notes');
  });

  it('satisfies review by EITHER a comment or a tracked change', () => {
    expect(intentsForManifest(manifest(['add-comment']))).toContain('review');
    expect(intentsForManifest(manifest(['tracked-change']))).toContain('review');
  });

  it('does not allow resolve-comment without a comment-reply actuation', () => {
    expect(intentsForManifest(manifest(['write-cells', 'add-comment']))).not.toContain(
      'resolve-comment',
    );
  });

  it('maps surface-specific generators to their actuation', () => {
    expect(intentsForManifest(manifest(['insert-slide']))).toContain('draft-slides');
    expect(intentsForManifest(manifest(['append-page']))).toContain('synthesize');
    expect(intentsForManifest(manifest(['post-message']))).toContain('meeting-notes');
  });
});
