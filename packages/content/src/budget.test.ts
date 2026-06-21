import { describe, it, expect } from 'vitest';
import { recommendStrategy, ContextBudget, DEFAULT_MAX_INLINE_TOKENS } from './budget.js';

describe('recommendStrategy', () => {
  it('inlines small, non-indexed content', () => {
    expect(recommendStrategy({ tokensEstimate: 500 })).toBe('inline');
  });

  it('references indexed content (regardless of size)', () => {
    expect(recommendStrategy({ tokensEstimate: 500, indexed: true })).toBe('reference');
    expect(recommendStrategy({ tokensEstimate: 999999, indexed: true })).toBe('reference');
  });

  it('uploads analytical data for code execution', () => {
    expect(recommendStrategy({ tokensEstimate: 200, analytical: true })).toBe(
      'upload-for-code-execution',
    );
  });

  it('uploads large non-indexed content for code execution', () => {
    expect(recommendStrategy({ tokensEstimate: DEFAULT_MAX_INLINE_TOKENS + 1 })).toBe(
      'upload-for-code-execution',
    );
  });
});

describe('ContextBudget', () => {
  it('reserves until the limit, then refuses', () => {
    const b = new ContextBudget(100);
    expect(b.tryAdd(60)).toBe(true);
    expect(b.tryAdd(60)).toBe(false); // would exceed
    expect(b.remaining).toBe(40);
    expect(b.fraction).toBeCloseTo(0.6);
  });
});
