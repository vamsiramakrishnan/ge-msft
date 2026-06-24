import { describe, it, expect } from 'vitest';
import type { Intent } from './intent.js';
import { QUICK_ACTIONS, QuickActionSchema, quickActionsForSurface } from './quick-actions.js';

describe('QUICK_ACTIONS catalog', () => {
  it('every quick action parses the schema', () => {
    for (const action of QUICK_ACTIONS) {
      expect(() => QuickActionSchema.parse(action)).not.toThrow();
    }
  });

  it('every action lists at least one surface and a non-empty prompt', () => {
    for (const action of QUICK_ACTIONS) {
      expect(action.surfaces.length).toBeGreaterThan(0);
      expect(action.prompt.length).toBeGreaterThan(0);
    }
  });

  it('ids are unique', () => {
    const ids = QUICK_ACTIONS.map((a) => a.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('quickActionsForSurface', () => {
  it('filters to actions offered on the surface', () => {
    const word = quickActionsForSurface('word');
    expect(word.length).toBeGreaterThan(0);
    expect(word.every((a) => a.surfaces.includes('word'))).toBe(true);
    // A Word action and a universal action are both present; an Excel-only one is not.
    expect(word.some((a) => a.id === 'review-policy')).toBe(true);
    expect(word.some((a) => a.id === 'summarize-this')).toBe(true);
    expect(word.some((a) => a.id === 'summarize-range')).toBe(false);
  });

  it('universal actions appear on every surface', () => {
    const surfaces = ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as const;
    for (const surface of surfaces) {
      const ids = quickActionsForSurface(surface).map((a) => a.id);
      expect(ids).toContain('summarize-this');
    }
  });

  it('drops actions whose intent is not in the allowed set (closure, ADR-0006)', () => {
    // Word offers a `review` action (review-policy) — exclude `review` and it disappears.
    const allowed: Intent[] = ['assist', 'regen-clause', 'resolve-comment'];
    const ids = quickActionsForSurface('word', allowed).map((a) => a.id);
    expect(ids).not.toContain('review-policy');
    expect(ids).not.toContain('find-unsupported');
    expect(ids).toContain('tighten'); // regen-clause is allowed
    expect(ids).toContain('summarize-this'); // assist is allowed
  });

  it('accepts a Set of allowed intents', () => {
    const ids = quickActionsForSurface('excel', new Set<Intent>(['assist'])).map((a) => a.id);
    expect(ids).toContain('summarize-range');
    expect(ids).not.toContain('find-anomalies'); // review excluded
  });
});
