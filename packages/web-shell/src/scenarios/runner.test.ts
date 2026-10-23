import { describe, expect, it, vi } from 'vitest';
import {
  checkScenario,
  geminiRefusal,
  runScenario,
  type Scenario,
  type ScenarioResult,
} from './runner.js';

const call = (status: number, reply = '') => ({ at: '', url: '', status, request: '', reply });

describe('source-grounded table outcomes', () => {
  const scenario: Scenario = {
    id: 'table-facts',
    surface: 'powerpoint',
    prompt: 'Tabulate these actions',
    seed: {},
    expect: {
      tableColumns: {
        Owner: { allowed: ['', 'TBD'] },
        Deadline: { allowed: ['', 'TBD', 'December'], includes: ['December'] },
      },
    },
  };
  const result = (table: string[][]): ScenarioResult => ({
    id: scenario.id,
    surface: scenario.surface,
    prompt: scenario.prompt,
    passed: false,
    failures: [],
    durationMs: 0,
    status: 'completed',
    applied: { 'add-table-slide': 1 },
    failedEffects: [],
    answer: '',
    errors: [],
    gates: [],
    steps: [],
    runs: [],
    transcript: [],
    document: { slides: [{ shapes: [{ table }] }] },
  });
  it('rejects invented cells despite a successful table write', () => {
    expect(
      checkScenario(
        scenario,
        result([
          ['Risk', 'Owner', 'Deadline'],
          ['MFA', 'Cyber team', 'Q4 2026'],
        ]),
      ),
    ).not.toEqual([]);
    expect(
      checkScenario(
        scenario,
        result([
          ['Risk', 'Owner', 'Deadline'],
          ['MFA', 'TBD', 'December'],
        ]),
      ),
    ).toEqual([]);
  });
  it('requires the table headers and retains source facts in the cells, not unrelated shapes', () => {
    expect(checkScenario(scenario, result([['Risk'], ['December']]))).not.toEqual([]);
    expect(
      checkScenario(
        scenario,
        result([
          ['Risk', 'Owner', 'Deadline'],
          ['MFA', 'TBD', 'TBD'],
        ]),
      ),
    ).not.toEqual([]);
  });

  it('rejects swapped owners and deadlines, and propagation of a partial deadline', () => {
    const paired: Scenario = {
      ...scenario,
      expect: {
        tableColumns: { Action: {}, Owner: {}, Deadline: {} },
        tableRows: [
          { Action: 'Audit', Owner: ['Mira'], Deadline: ['November'] },
          { Action: 'Review', Owner: ['TBD'], Deadline: ['TBD'] },
        ],
      },
    };
    expect(
      checkScenario(
        paired,
        result([
          ['Action', 'Owner', 'Deadline'],
          ['Audit', 'TBD', 'TBD'],
          ['Review', 'Mira', 'November'],
        ]),
      ),
    ).not.toEqual([]);
    expect(
      checkScenario(
        paired,
        result([
          ['Action', 'Owner', 'Deadline'],
          ['Audit', 'Mira', 'November'],
          ['Review', 'TBD', 'November'],
        ]),
      ),
    ).not.toEqual([]);
    expect(
      checkScenario(
        paired,
        result([
          ['Action', 'Owner', 'Deadline'],
          ['Audit', 'Mira', 'November'],
          ['Review', 'TBD', 'TBD'],
        ]),
      ),
    ).toEqual([]);
  });

  it('requires two populated sample rows instead of accepting only headers or blank cells', () => {
    const examples: Scenario = {
      ...scenario,
      expect: {
        tableColumns: { Owner: { nonEmpty: true }, Deadline: { nonEmpty: true } },
        tableRowCount: 2,
      },
    };
    expect(checkScenario(examples, result([['Owner', 'Deadline']]))).not.toEqual([]);
    expect(
      checkScenario(
        examples,
        result([
          ['Owner', 'Deadline'],
          ['', ''],
          ['', ''],
        ]),
      ),
    ).not.toEqual([]);
    expect(
      checkScenario(
        examples,
        result([
          ['Owner', 'Deadline'],
          ['Fictional A', 'July'],
          ['Fictional B', 'August'],
        ]),
      ),
    ).toEqual([]);
  });
});

describe('slide insertion outcomes', () => {
  const scenario: Scenario = {
    id: 'placement',
    surface: 'powerpoint',
    prompt: 'Add a slide between the first and second slides',
    seed: { slides: [{ id: 'a' }, { id: 'b' }, { id: 'c' }] },
    expect: { applied: { 'insert-slide': 1 }, insertedSlidePositions: [2] },
  };
  const result = (ids: string[]): ScenarioResult => ({
    id: scenario.id,
    surface: scenario.surface,
    prompt: scenario.prompt,
    passed: false,
    failures: [],
    durationMs: 0,
    status: 'completed',
    applied: { 'insert-slide': 1 },
    failedEffects: [],
    answer: '',
    errors: [],
    gates: [],
    steps: [],
    runs: [],
    transcript: [],
    document: { slides: ids.map((id) => ({ id })) },
  });

  it('rejects append despite a successful insertion receipt', () => {
    expect(checkScenario(scenario, result(['a', 'b', 'c', 'new']))).toEqual([
      'new slides are at [4], expected [2]',
    ]);
    expect(checkScenario(scenario, result(['a', 'new', 'b', 'c']))).toEqual([]);
  });

  it('checks multiple insertions and fails closed when slide IDs are missing', () => {
    const multiple = { ...scenario, expect: { insertedSlidePositions: [1, 4] } };
    expect(checkScenario(multiple, result(['x', 'a', 'b', 'y', 'c']))).toEqual([]);
    expect(checkScenario(multiple, result(['a', 'b', 'c', 'x', 'y']))).not.toEqual([]);
    expect(checkScenario({ ...scenario, seed: { slides: [{}] } }, result(['new']))).toEqual([
      'slide positions require seed and result slides with stable IDs',
    ]);
  });

  it('rejects deleting or reordering original slides even if the new position is correct', () => {
    expect(checkScenario(scenario, result(['b', 'new', 'c']))).toContain(
      'original slides were removed or reordered',
    );
    expect(checkScenario(scenario, result(['c', 'new', 'a', 'b']))).toContain(
      'original slides were removed or reordered',
    );
  });

  it('keeps the seed immutable while checking an actual local slide insertion', async () => {
    const input: Scenario = {
      ...scenario,
      prompt: '```cmd\nslide "Added" "One bullet" at=2\n```',
      seed: {
        slides: [
          { id: 'a', shapes: [{ text: 'First' }] },
          { id: 'b', shapes: [{ text: 'Second' }] },
        ],
      },
    };
    const before = structuredClone(input.seed);
    // Placeholder config so a clean checkout (no packages/web-shell/.env) runs this offline test.
    const guid = '00000000-0000-0000-0000-000000000000';
    for (const [key, value] of Object.entries({
      VITE_GCP_PROJECT: 'test-project',
      VITE_GCP_LOCATION: 'global',
      VITE_GE_ENGINE: 'test-engine',
      VITE_WIF_POOL_ID: 'test-pool',
      VITE_WIF_PROVIDER_ID: 'test-provider',
      VITE_ENTRA_TENANT_ID: guid,
      VITE_ENTRA_CLIENT_ID: guid,
    }))
      vi.stubEnv(key, process.env[key] || value);
    const network = vi
      .spyOn(globalThis, 'fetch')
      .mockRejectedValue(new Error('unexpected network'));
    try {
      const actual = await runScenario(input);
      expect(input.seed).toEqual(before);
      expect(actual.failures).toEqual([]);
      expect(actual.applied['insert-slide']).toBe(1);
      expect(actual.transcript).toEqual([]);
    } finally {
      network.mockRestore();
      vi.unstubAllEnvs();
    }
  }, 12000);
});

describe('geminiRefusal (WS5 T3)', () => {
  it('names the refusal when no call got an answer', () => {
    expect(geminiRefusal([])).toBe('no Gemini call was made');
    expect(geminiRefusal([call(429), call(429)])).toBe('Gemini refused every call (HTTP 429)');
    expect(geminiRefusal([call(401), call(503)])).toBe('Gemini refused every call (HTTP 401, 503)');
    expect(geminiRefusal([call(200, '')])).toBe('Gemini answered nothing (HTTP 200, empty)');
  });

  it('is undefined once any call got an answer: the run is a real result', () => {
    expect(geminiRefusal([call(200, '```plan\nintent ask\n```'), call(429)])).toBeUndefined();
  });

  it('keeps a local command result even though it does not call Gemini', () => {
    expect(geminiRefusal([], true)).toBeUndefined();
    expect(geminiRefusal([call(429)], true)).toBe('Gemini refused every call (HTTP 429)');
  });

  it('keeps a failed HTTP 200 stream as a failed run, not a refusal', () => {
    expect(geminiRefusal([{ ...call(200), bodyState: 'failed' }])).toBeUndefined();
    expect(geminiRefusal([{ ...call(429), bodyState: 'failed' }])).toBe(
      'Gemini refused every call (HTTP 429)',
    );
  });

  it('keeps an active HTTP 200 stream at the scenario deadline as an executed run', () => {
    expect(geminiRefusal([{ ...call(200), bodyState: 'streaming' }])).toBeUndefined();
    expect(geminiRefusal([{ ...call(200), bodyState: 'complete' }])).toBe(
      'Gemini answered nothing (HTTP 200, empty)',
    );
  });
});
