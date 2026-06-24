import { describe, it, expect } from 'vitest';
import { CommandPlanSchema, extractPlanBlock, parsePlanBlock } from './command-plan.js';

/** The self-test sample from parse_plan.py (an unknown keyword line included on purpose). */
const SAMPLE = `**thought** mapping the request to a plan
\`\`\`plan
intent   review
surface  word
scope    §4-6
ground   "Vendor Risk Policy v4"
step     flag clauses in §4-6 that breach APRA CPS 234
step     rewrite the SLA availability figure to 99.9% as a tracked change
exclude  the indemnity clause — leave unchanged
confidence high
prioritise nonsense
\`\`\``;

describe('parsePlanBlock', () => {
  it('parses the self-test sample into a structured plan', () => {
    const { plan, errors, needsClarification } = parsePlanBlock(SAMPLE);
    expect(plan).not.toBeNull();
    expect(plan!.intent).toBe('review');
    expect(plan!.surface).toBe('word');
    expect(plan!.scope).toBe('§4-6');
    expect(plan!.ground).toEqual(['Vendor Risk Policy v4']); // quotes stripped
    expect(plan!.steps).toHaveLength(2);
    expect(plan!.excludes).toEqual(['the indemnity clause — leave unchanged']);
    expect(plan!.confidence).toBe('high');
    expect(needsClarification).toBe(false);
    // The bogus `prioritise` line is reported, not silently dropped.
    expect(errors.some((e) => e.includes("unknown plan keyword 'prioritise'"))).toBe(true);
    // The structured plan still validates against the schema.
    expect(() => CommandPlanSchema.parse(plan)).not.toThrow();
  });

  it('parses a minimal good plan with no errors', () => {
    const text = '```plan\nintent assist\nsurface excel\nstep summarize the range\n```';
    const { plan, errors, needsClarification } = parsePlanBlock(text);
    expect(errors).toEqual([]);
    expect(needsClarification).toBe(false);
    expect(plan).toEqual({
      intent: 'assist',
      surface: 'excel',
      ground: [],
      steps: ['summarize the range'],
      excludes: [],
      clarify: [],
    });
  });

  it('flags an unknown keyword with a did-you-mean', () => {
    const text = '```plan\nintent review\nsurface word\nstep do it\nground2 nope\n```';
    const { errors } = parsePlanBlock(text);
    const unknown = errors.find((e) => e.includes("unknown plan keyword 'ground2'"));
    expect(unknown).toBeDefined();
    expect(unknown).toContain("did you mean 'ground'");
  });

  it('sets needsClarification on a clarify-only plan and does not require a step', () => {
    const text =
      '```plan\nintent review\nsurface word\nclarify which sections should I review?\n```';
    const { plan, errors, needsClarification } = parsePlanBlock(text);
    expect(needsClarification).toBe(true);
    // A clarify substitutes for a step → no "needs at least one step" error.
    expect(errors).toEqual([]);
    expect(plan!.clarify).toEqual(['which sections should I review?']);
    expect(plan!.steps).toEqual([]);
  });

  it('reports missing intent and surface, returning a null plan', () => {
    const text = '```plan\nstep do something\n```';
    const { plan, errors } = parsePlanBlock(text);
    expect(plan).toBeNull();
    expect(errors).toContain("plan is missing 'intent'");
    expect(errors).toContain("plan is missing 'surface'");
  });

  it('rejects an invalid intent and surface value', () => {
    const text = '```plan\nintent frobnicate\nsurface mars\nstep go\n```';
    const { plan, errors } = parsePlanBlock(text);
    expect(plan).toBeNull();
    expect(errors.some((e) => e.includes("unknown intent 'frobnicate'"))).toBe(true);
    expect(errors.some((e) => e.includes("unknown surface 'mars'"))).toBe(true);
  });

  it('returns a null plan and no errors when there is no plan fence (re-prompt)', () => {
    const { plan, errors, needsClarification } = parsePlanBlock('just some prose, no plan here');
    expect(plan).toBeNull();
    expect(errors).toEqual([]);
    expect(needsClarification).toBe(false);
  });
});

describe('extractPlanBlock', () => {
  it('tolerates an unclosed fence', () => {
    const inner = extractPlanBlock('```plan\nintent assist\nsurface excel\nstep go');
    expect(inner).toBe('intent assist\nsurface excel\nstep go');
  });

  it('returns null when there is no fence', () => {
    expect(extractPlanBlock('no fence at all')).toBeNull();
  });
});
