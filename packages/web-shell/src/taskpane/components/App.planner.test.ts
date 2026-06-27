import { describe, it, expect } from 'vitest';
import { isComplexInstruction, isActuating } from './App.js';

/**
 * The §F routing heuristic: a complex free-text instruction earns the planner-confirm front door; a
 * short single-shot one (or a chat verb) does not. `isActuating` decides chat-vs-gate; together they
 * gate whether dispatch routes through `proposePlan`.
 */
describe('isComplexInstruction (planner-confirm front door, EXPERIENCE.md §F)', () => {
  it('treats a short single-shot instruction as simple', () => {
    expect(isComplexInstruction('make it formal')).toBe(false);
    expect(isComplexInstruction('tighten this')).toBe(false);
    expect(isComplexInstruction('')).toBe(false);
  });

  it('treats a constraint/exclusion instruction as complex', () => {
    expect(isComplexInstruction('rewrite the SLA to 99.9% but leave the indemnity clause')).toBe(
      true,
    );
    expect(isComplexInstruction('tighten §4, and flag anything below policy')).toBe(true);
    expect(isComplexInstruction('summarize only the risk section')).toBe(true);
  });

  it('treats a long instruction as complex even without a marker word', () => {
    expect(
      isComplexInstruction(
        'rewrite this paragraph so it reads as a crisp executive overview please',
      ),
    ).toBe(true);
  });
});

describe('isActuating', () => {
  it('is true exactly for the write/annotation verbs', () => {
    for (const v of ['rewrite', 'review', 'draft', 'notes'] as const)
      expect(isActuating(v)).toBe(true);
    for (const v of ['ask', 'summarize', 'explain'] as const) expect(isActuating(v)).toBe(false);
    expect(isActuating(undefined)).toBe(false);
  });
});
