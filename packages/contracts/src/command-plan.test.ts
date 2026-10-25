import { describe, it, expect } from 'vitest';
import {
  commandIntentText,
  confirmedPlanChanges,
  planChanges,
  editIntentFor,
  planDescribesChange,
  CONFIRMED_PLAN_CLOSE,
  CONFIRMED_PLAN_OPEN,
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

  it('accepts visualize as a first-class planner intent', () => {
    const text = `\`\`\`plan
intent visualize
surface excel
scope selection
step derive a chart-ready summary table before creating a native chart
\`\`\``;
    const { plan, errors } = parsePlanBlock(text);
    expect(errors).toEqual([]);
    expect(plan?.intent).toBe('visualize');
    expect(plan?.steps).toEqual([
      'derive a chart-ready summary table before creating a native chart',
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
    expect(prompt).toContain('capability-shaped steps');
    expect(prompt).toContain('one rectangular grid/table materialization step');
  });

  it('reads as a task with no role or override wording, and allows the engine follow-up block', () => {
    // 2026-10-19: "emit only the plan" against the engine's "always add suggested_questions"
    // made Gemini 3.8 reason about overriding system instructions; Model Armor blocked it.
    const prompt = renderPlanPrompt('excel');
    expect(prompt).not.toMatch(/You are|Do NOT|emit only|block only/);
    expect(prompt).toContain('A suggested_questions block after the plan is fine');
  });

  it('parses the plan when the engine adds its follow-up block after it', () => {
    const reply =
      '```plan\nintent rewrite\nsurface excel\nstep add a growth column\n```\n\n```suggested_questions\n["Add a chart?"]\n```';
    expect(parsePlanBlock(reply).plan).toMatchObject({ intent: 'rewrite', surface: 'excel' });
  });
});

describe('commandIntentText', () => {
  it('keeps only the request, steps and exclusions of a confirmed-plan task', () => {
    const task = [
      'Execute this user-confirmed plan in the open Microsoft 365 surface.',
      'Treat the plan as approved intent only: read live host content before any write.',
      '',
      CONFIRMED_PLAN_OPEN,
      'original_request: Change the title on slide 1',
      'intent: draft',
      'surface: powerpoint',
      'step 1: set the title shape text',
      'exclude: other slides',
      'confidence: high',
      CONFIRMED_PLAN_CLOSE,
    ].join('\n');
    expect(commandIntentText(task)).toBe(
      'Change the title on slide 1\nset the title shape text\nother slides',
    );
  });

  it('returns any other task unchanged', () => {
    expect(commandIntentText('reply to the email accepting the meeting')).toBe(
      'reply to the email accepting the meeting',
    );
    expect(commandIntentText(`${CONFIRMED_PLAN_OPEN} unterminated`)).toBe(
      `${CONFIRMED_PLAN_OPEN} unterminated`,
    );
  });
});

describe('live write vocabulary', () => {
  it('plans only the writes the surface offers, and says no to the rest', () => {
    const prompt = renderPlanPrompt('excel', ['ask', 'rewrite'], ['Write cells', 'Format cells']);
    expect(prompt).toContain('Changes this add-in can make: Write cells; Format cells.');
    expect(prompt).toContain('never approximate it with other changes');
    // The static vocabulary named unshipped features (pivot, sort) and invited plans for them.
    expect(prompt).not.toContain('pivot');
  });
});

describe('intent effects (fix G)', () => {
  it('describes only the offered intents by their effect', () => {
    const prompt = renderPlanPrompt('excel', ['ask', 'rewrite', 'visualize']);
    expect(prompt).toContain('rewrite   change existing content: text, values, formulas');
    expect(prompt).toContain(
      'A request that changes the document never uses ask, summarize or explain.',
    );
    expect(prompt).not.toContain('draft     create or change content');
  });

  it('recognises change steps and the edit intent of each surface', () => {
    expect(planDescribesChange({ steps: ['add excel formula in G12 to sum range G2:G11'] })).toBe(
      true,
    );
    // Shaping the answer is not a document change (live 2026-10-09).
    expect(
      planDescribesChange({ steps: ['Format the extracted action items as a clear list'] }),
    ).toBe(false);
    expect(
      planDescribesChange({
        steps: ['extract and format the list of identified supplier obligations'],
      }),
    ).toBe(false);
    expect(planDescribesChange({ steps: ['format the header row bold'] })).toBe(true);
    expect(planDescribesChange({ steps: ['format B2:E11 as currency'] })).toBe(true);
    // Analysis wording is an answer, not formatting (live 2026-10-12: "apa yang tampak tidak biasa").
    expect(
      planDescribesChange({ steps: ['Highlight any unusual patterns or outliers in the data'] }),
    ).toBe(false);
    expect(planDescribesChange({ steps: ['highlight cells above 100 in green'] })).toBe(true);
    expect(planDescribesChange({ steps: ['highlight the outliers in D2:D10 in red'] })).toBe(true);
    // Security review 2026-10-12: formatting without a cell reference is still a change.
    expect(planDescribesChange({ steps: ['highlight the outlier cells in red'] })).toBe(true);
    expect(planDescribesChange({ steps: ['highlight unusual values in yellow'] })).toBe(true);
    expect(planDescribesChange({ steps: ['flag the outlier rows'] })).toBe(true);
    // A real change in the same step is still a change (security review 2026-10-12).
    expect(
      planDescribesChange({
        steps: ['Compile the findings and add them as a comment on paragraph 3'],
      }),
    ).toBe(true);
    expect(
      planDescribesChange({ steps: ['Format the identified totals in D2:D10 as currency'] }),
    ).toBe(true);
    expect(
      planDescribesChange({ steps: ['Organize the extracted rows and write them to Sheet2'] }),
    ).toBe(true);
    expect(planDescribesChange({ steps: ['identify the region with the highest total'] })).toBe(
      false,
    );
    expect(editIntentFor('excel')).toBe('rewrite');
    expect(editIntentFor('word')).toBe('rewrite');
    expect(editIntentFor('powerpoint')).toBe('draft');
    expect(editIntentFor('outlook')).toBe('draft');
    expect(editIntentFor('teams')).toBe('notes');
    expect(editIntentFor('not-a-surface' as never)).toBeUndefined();
  });
});

describe('changes: the planner states the effect (WS1, 2026-10-20)', () => {
  const plan = (body: string) =>
    parsePlanBlock(`\`\`\`plan\nintent   ask\nsurface  excel\n${body}\n\`\`\``);

  it('parses changes yes|no into a boolean', () => {
    expect(plan('changes  yes\nstep     leave a note on the lowest region').plan?.changes).toBe(
      true,
    );
    expect(plan('changes  No\nstep     filter rows where Region is Kalimantan').plan?.changes).toBe(
      false,
    );
    expect(plan('step     answer').plan?.changes).toBeUndefined();
  });

  it('rejects any other value', () => {
    const { errors } = plan('changes  maybe\nstep     answer');
    expect(errors.join(' ')).toMatch(/changes must be yes\|no/);
  });

  it('trusts the stated effect over the step keywords, and falls back only without it', () => {
    // Argon mx-xl-read-14: an analysis step with "filter" + "rows" is not a change.
    const analysis = { steps: ['filter rows where Region is Kalimantan and compare totals'] };
    expect(planChanges({ ...analysis, changes: false })).toEqual({ changes: false, stated: true });
    expect(planChanges(analysis)).toEqual({ changes: true, stated: false });
    // u-xl-note-on-lowest: a question that also leaves a note is a change.
    expect(planChanges({ steps: ['find the lowest region'], changes: true })).toEqual({
      changes: true,
      stated: true,
    });
  });

  it('reads the stated effect back from a confirmed-plan task', () => {
    const task = (line: string) =>
      [
        CONFIRMED_PLAN_OPEN,
        'original_request: changes: yes please',
        'intent: ask',
        line,
        'step 1: x',
        CONFIRMED_PLAN_CLOSE,
      ].join('\n');
    expect(confirmedPlanChanges(task('changes: no'))).toBe(false);
    expect(confirmedPlanChanges(task('changes: yes'))).toBe(true);
    // The user's own words never count: only the plan's line after the request.
    expect(confirmedPlanChanges(task('confidence: high'))).toBeUndefined();
    expect(confirmedPlanChanges('changes: yes')).toBeUndefined();
  });

  it('asks the planner for it and defines it by effect', () => {
    const prompt = renderPlanPrompt('excel');
    expect(prompt).toContain('changes  <yes|no>');
    expect(prompt).toMatch(/question or analysis is no/i);
    expect(prompt).toMatch(/don't change anything/i);
  });
});
