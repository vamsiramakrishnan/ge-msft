import { describe, it, expect } from 'vitest';
import {
  CommandPlanSchema,
  derivePlanContextStrategy,
  describePlanContextHints,
  extractPlanBlock,
  parsePlanBlock,
  renderPlanPrompt,
} from './command-plan.js';

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
    // A free-text scope degrades to a section heading ref.
    expect(plan!.scope).toEqual({ kind: 'section', ref: '§4-6' });
    // A named source (quotes stripped) becomes a `document` ground token.
    expect(plan!.ground).toEqual([{ kind: 'document', ref: 'Vendor Risk Policy v4' }]);
    expect(plan!.context).toEqual([]);
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
    const text = '```plan\nintent ask\nsurface excel\nstep summarize the range\n```';
    const { plan, errors, needsClarification } = parsePlanBlock(text);
    expect(errors).toEqual([]);
    expect(needsClarification).toBe(false);
    expect(plan).toEqual({
      intent: 'ask',
      surface: 'excel',
      ground: [],
      context: [],
      steps: ['summarize the range'],
      excludes: [],
      clarify: [],
    });
  });

  it('parses the live plain plan sentinel shape without a fence', () => {
    const text = `plan
intent   rewrite
surface  excel
scope    document
step     Generate a mock weekly schedule for a Google SWE based in Sunnyvale
exclude  overwrite existing headers
confidence high`;
    expect(extractPlanBlock(text)).toBe(text);
    const { plan, errors, needsClarification } = parsePlanBlock(text);
    expect(errors).toEqual([]);
    expect(needsClarification).toBe(false);
    expect(plan?.intent).toBe('rewrite');
    expect(plan?.scope).toEqual({ kind: 'document' });
    expect(plan?.steps).toEqual([
      'Generate a mock weekly schedule for a Google SWE based in Sunnyvale',
    ]);
    expect(plan?.excludes).toEqual(['overwrite existing headers']);
  });

  it('parses a bare ground token and a function-form scope', () => {
    const text =
      '```plan\nintent rewrite\nsurface word\nscope section(§4)\nground unit\nground this\nstep tighten\n```';
    const { plan } = parsePlanBlock(text);
    expect(plan!.scope).toEqual({ kind: 'section', ref: '§4' });
    expect(plan!.ground).toEqual([{ kind: 'unit' }, { kind: 'this' }]);
  });

  it('parses repeatable context hints for an analytical Excel plan', () => {
    const text = `\`\`\`plan
intent draft
surface excel
scope document
ground this
context analytical
context full-scope
context upload-preferred
context code-execution-preferred
step create a chart-ready risk summary table
\`\`\``;
    const { plan, errors } = parsePlanBlock(text);
    expect(errors).toEqual([]);
    expect(plan?.context).toEqual([
      'analytical',
      'full-scope',
      'upload-preferred',
      'code-execution-preferred',
    ]);
  });

  it('derives one shared progressive-disclosure strategy from context hints', () => {
    const strategy = derivePlanContextStrategy([
      'analytical',
      'full-scope',
      'upload-preferred',
      'code-execution-preferred',
      'analytical',
    ]);
    expect(strategy.scope).toBe('whole-artifact');
    expect(strategy.transfer).toBe('upload-candidate');
    expect(strategy.analysis).toBe('code-execution-candidate');
    expect(strategy.hints.map((h) => h.label)).toEqual([
      'Analytical',
      'Full scope',
      'File upload preferred',
      'Code execution preferred',
    ]);
    expect(describePlanContextHints(undefined)).toEqual([]);
  });

  it('accepts context hints across every surface without broadening planner intents', () => {
    const surfaces = ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as const;
    for (const surface of surfaces) {
      const text = `\`\`\`plan
intent summarize
surface ${surface}
scope document
context full-scope
context incremental
step summarize the full open artifact
\`\`\``;
      const { plan, errors } = parsePlanBlock(text);
      expect(errors).toEqual([]);
      expect(plan?.surface).toBe(surface);
      expect(plan?.context).toEqual(['full-scope', 'incremental']);
    }
  });

  it('rejects unknown context hints instead of passing through free text', () => {
    const text = '```plan\nintent ask\nsurface excel\ncontext run-python-now\nstep inspect\n```';
    const { plan, errors } = parsePlanBlock(text);
    expect(plan?.context).toEqual([]);
    expect(errors.some((e) => e.includes("unknown context hint 'run-python-now'"))).toBe(true);
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

describe('renderPlanPrompt', () => {
  it('advertises context hints without granting execution authority', () => {
    const prompt = renderPlanPrompt('excel', ['ask', 'draft']);
    expect(prompt).toContain('context  <incremental|inline-preferred');
    expect(prompt).toContain('context is only a context-construction hint');
    expect(prompt).toContain('never grants upload/code/write authority');
    expect(prompt).toContain('intent   <ask | draft>');
  });
});
