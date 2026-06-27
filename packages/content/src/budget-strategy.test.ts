import { describe, it, expect } from 'vitest';
import {
  recommendStrategy,
  ContextBudget,
  DEFAULT_MAX_INLINE_TOKENS,
  DEFAULT_CONTEXT_BUDGET_TOKENS,
} from './budget.js';
import { contextualizeChunk } from './contextualize.js';
import type { Chunk } from './model.js';

/**
 * Decision-table coverage of recommendStrategy and the running ContextBudget accumulator,
 * plus the contextualize header for the no-breadcrumb case.
 */
describe('recommendStrategy', () => {
  it('always uploads for code execution when analytical, regardless of size or index', () => {
    expect(recommendStrategy({ tokensEstimate: 10, analytical: true })).toBe(
      'upload-for-code-execution',
    );
    // analytical wins even over an indexed source that would otherwise reference.
    expect(recommendStrategy({ tokensEstimate: 999999, analytical: true, indexed: true })).toBe(
      'upload-for-code-execution',
    );
  });

  it('references an indexed source that exceeds the inline budget', () => {
    expect(
      recommendStrategy({ tokensEstimate: DEFAULT_MAX_INLINE_TOKENS + 1, indexed: true }),
    ).toBe('reference');
  });

  it('references an indexed source even when it fits the inline budget (ACL-preserving)', () => {
    expect(recommendStrategy({ tokensEstimate: 100, indexed: true })).toBe('reference');
  });

  it('inlines a small non-indexed source', () => {
    expect(recommendStrategy({ tokensEstimate: 100 })).toBe('inline');
    // Exactly at the boundary still inlines (<=).
    expect(recommendStrategy({ tokensEstimate: DEFAULT_MAX_INLINE_TOKENS })).toBe('inline');
  });

  it('uploads a large non-indexed source that overflows the inline budget', () => {
    expect(recommendStrategy({ tokensEstimate: DEFAULT_MAX_INLINE_TOKENS + 1 })).toBe(
      'upload-for-code-execution',
    );
  });

  it('honours a custom maxInlineTokens override', () => {
    // 50 tokens is over a tiny custom budget of 10 -> non-indexed overflow -> upload.
    expect(recommendStrategy({ tokensEstimate: 50, maxInlineTokens: 10 })).toBe(
      'upload-for-code-execution',
    );
    // ...but under it -> inline.
    expect(recommendStrategy({ tokensEstimate: 5, maxInlineTokens: 10 })).toBe('inline');
  });
});

describe('ContextBudget', () => {
  it('reserves tokens while they fit and rejects an item that would overflow', () => {
    const b = new ContextBudget(100);
    expect(b.tryAdd(60)).toBe(true);
    expect(b.remaining).toBe(40);
    // 50 would push used to 110 > 100 -> rejected, state unchanged.
    expect(b.tryAdd(50)).toBe(false);
    expect(b.remaining).toBe(40);
    // The exact remaining amount still fits.
    expect(b.tryAdd(40)).toBe(true);
    expect(b.remaining).toBe(0);
  });

  it('reports fraction used and clamps to 1', () => {
    const b = new ContextBudget(200);
    expect(b.fraction).toBe(0);
    b.tryAdd(50);
    expect(b.fraction).toBeCloseTo(0.25);
  });

  it('treats a zero limit as full (fraction 1, nothing fits)', () => {
    const b = new ContextBudget(0);
    expect(b.fraction).toBe(1);
    expect(b.tryAdd(1)).toBe(false);
    expect(b.remaining).toBe(0);
  });

  it('defaults to the documented total budget', () => {
    const b = new ContextBudget();
    expect(b.remaining).toBe(DEFAULT_CONTEXT_BUDGET_TOKENS);
  });
});

describe('contextualizeChunk', () => {
  const baseChunk = (over: Partial<Chunk['meta']>): Chunk => ({
    id: 's#0',
    index: 0,
    text: 'body text',
    meta: {
      sourceId: 's',
      sectionPath: [],
      kinds: ['paragraph'],
      tokensEstimate: 2,
      anchor: { matchText: 'body text' },
      ...over,
    },
  });

  it('returns the bare chunk text when there is no source title and no section path', () => {
    expect(contextualizeChunk(baseChunk({}))).toBe('body text');
  });

  it('prepends only the source title when there is no section path', () => {
    expect(contextualizeChunk(baseChunk({ sourceTitle: 'MSA' }))).toBe('[MSA]\nbody text');
  });

  it('joins source title and section breadcrumb with the › separator', () => {
    const out = contextualizeChunk(
      baseChunk({ sourceTitle: 'MSA', sectionPath: ['5. Levels', '5.1 Avail'] }),
    );
    expect(out).toBe('[MSA › 5. Levels › 5.1 Avail]\nbody text');
  });

  it('uses only the section path when there is no source title', () => {
    expect(contextualizeChunk(baseChunk({ sectionPath: ['Intro'] }))).toBe('[Intro]\nbody text');
  });
});
