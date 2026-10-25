import { describe, expect, it, vi } from 'vitest';
import { contentHash, type ActuationRequest, type ActuationResult } from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { NativeTraceStore, observeNativeBatch } from './native-trace.js';
import {
  runOfficeBenchmarkComparison,
  type OfficeBenchmarkComparisonInput,
  type OfficeBenchmarkLegInput,
} from './office-benchmark.js';

const hash = (value: string) => contentHash({ value });
const passedChecks = {
  targetIntegrity: 'passed',
  approvalEnforced: 'passed',
  completionTruth: 'passed',
  freshnessRefusal: 'passed',
} as const;
function fixture(store: NativeTraceStore, syncs: number, id: string) {
  let text = 'old';
  let source = 'private-source-never-export';
  let outcome = 'verified';
  let platform = 'test-only';
  let key = id;
  let probeFailure = false;
  let checks: { [K in keyof typeof passedChecks]: 'passed' | 'failed' | 'unavailable' } = {
    ...passedChecks,
  };
  let foreign = false;
  const meter = { requests: 0, inputBytes: 0, complete: true };
  const approval = vi.fn(async () => true);
  const fetch = vi.fn(async (_url: unknown, init?: RequestInit) => {
    meter.requests++;
    meter.inputBytes += new TextEncoder().encode(String(init?.body ?? '')).byteLength;
    const reply = '```cmd\nsuggest "old" => "new"\nfinish when=verified\n```';
    return new Response(
      JSON.stringify([
        {
          answer: {
            state: 'SUCCEEDED',
            replies: [{ groundedContent: { content: { text: reply } } }],
          },
        },
      ]),
    );
  });
  const actuate = vi.fn(async (request: ActuationRequest): Promise<ActuationResult> => {
    await observeNativeBatch(
      'word',
      { sync: async () => undefined },
      async (ctx) => {
        for (let i = 0; i < syncs; i++) await ctx.sync();
        text = request.params.text ?? text;
      },
      store,
    );
    if (foreign)
      store.record({
        id: crypto.randomUUID(),
        surface: 'word',
        stage: 'dispatch',
        kind: 'tracked-change',
        changeId: 'unrelated-change',
        syncs: 0,
        syncMs: 0,
        durationMs: 1,
        outcome: 'completed',
      });
    return {
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      ...(outcome === 'verified'
        ? { verification: { status: 'verified' as const } }
        : outcome === 'uncertain'
          ? { recoveryPending: true }
          : {}),
    };
  });
  const bridge: DocBridge = {
    surface: 'word',
    documentKey: () => key,
    getHostProfile: () => ({
      surface: 'word',
      platform,
      build: 'test-only-build',
      mode: 'write',
      requirements: { 'WordApi:1.4': true },
    }),
    getCapabilities: () => ({
      surface: 'word',
      contextKinds: [],
      actuations: [
        { kind: 'tracked-change', surface: 'word', title: 'Tracked change', reversible: true },
      ],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    actuate,
  };
  const client = new StreamAssistClient(
    { getAccessToken: async () => 'private-token-never-export' },
    {
      assistant: { project: 'p', location: 'eu', engine: 'e' },
      identity: 'private-identity-never-export',
    },
    fetch as never,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    autoAttach: [],
    context: { docState: false, lazyRead: false },
  });
  const leg: OfficeBenchmarkLegInput = {
    session,
    bridge,
    disposable: true,
    execution: { kind: 'program', program: 'suggest "old" => "new"\nfinish when=verified' },
    expectedEffects: [{ kind: 'tracked-change', completion: 'verified' }],
    options: { approvePlan: approval },
    inputMeter: { read: () => ({ ...meter }) },
    probe: {
      id: 'word-text-v1',
      evaluatorId: 'exact-text-v1',
      input: async () => ({
        scopeHash: await hash(key),
        sourceHash: await hash(source),
        artifactHash: await hash(text),
      }),
      outcome: async () => {
        if (probeFailure) throw new Error('private-host-error-never-export');
        return {
          scopeHash: await hash(key),
          sourceHash: await hash(source),
          artifactHash: await hash(text),
          checks,
        };
      },
    },
  };
  return {
    leg,
    approval,
    actuate,
    fetch,
    meter,
    setSource: (v: string) => {
      source = v;
    },
    setOutcome: (v: string) => {
      outcome = v;
    },
    setPlatform: (v: string) => {
      platform = v;
    },
    setKey: (v: string) => {
      key = v;
    },
    failProbe: () => {
      probeFailure = true;
    },
    setChecks: (v: typeof checks) => {
      checks = v;
    },
    contaminate: () => {
      foreign = true;
    },
    setText: (v: string) => {
      text = v;
    },
  };
}
function pair() {
  const store = new NativeTraceStore();
  const baseline = fixture(store, 4, 'baseline-document');
  const candidate = fixture(store, 2, 'candidate-document');
  const input: OfficeBenchmarkComparisonInput = {
    commit: 'a'.repeat(40),
    workflowId: 'contract-review',
    profileId: 'word-test',
    hostScope: 'simulated',
    providerScope: 'none',
    confirmExclusiveMeasurement: true,
    baseline: baseline.leg,
    candidate: candidate.leg,
  };
  return { store, baseline, candidate, input };
}
describe('measured Office comparisons use real session execution and labelled simulated hosts', () => {
  it('measures four versus two actual awaited syncs with identical scoped artifacts and separate approvals', async () => {
    const f = pair();
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('passed');
    expect(report.hostScope).toBe('simulated');
    expect(report.baseline.nativeSyncs).toBe(4);
    expect(report.candidate.nativeSyncs).toBe(2);
    expect(report.measurements).toEqual({ syncReduction: 0.5, inferenceInputReduction: null });
    expect(report.baseline.inferenceInputBytes).toBe(0);
    expect(f.baseline.approval).toHaveBeenCalledOnce();
    expect(f.candidate.approval).toHaveBeenCalledOnce();
    expect(f.baseline.actuate).toHaveBeenCalledOnce();
    expect(f.candidate.actuate).toHaveBeenCalledOnce();
    expect(f.baseline.fetch).not.toHaveBeenCalled();
    const exported = JSON.stringify(report);
    for (const privateText of [
      'private-source',
      'private-token',
      'private-identity',
      'baseline-document',
      'candidate-document',
      'suggest',
      'old',
    ])
      expect(exported).not.toContain(privateText);
  });
  it('cannot promote a simulator profile to native acceptance', async () => {
    const f = pair();
    f.input.hostScope = 'native';
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('unavailable');
    expect(report.issueCodes).toContain('unavailable-host');
    expect(f.baseline.actuate).not.toHaveBeenCalled();
    expect(f.candidate.actuate).not.toHaveBeenCalled();
    expect(report.measurements.syncReduction).toBeNull();
  });
  it('does not turn zero baseline syncs into a reduction claim', async () => {
    const store = new NativeTraceStore();
    const a = fixture(store, 0, 'a');
    const b = fixture(store, 0, 'b');
    const input = { ...pair().input, baseline: a.leg, candidate: b.leg };
    const report = await runOfficeBenchmarkComparison(input, store);
    expect(report.status).toBe('passed');
    expect(report.measurements.syncReduction).toBeNull();
  });
  it('retains refusal and uncertain outcomes as failed comparisons with no reduction claim', async () => {
    const f = pair();
    f.candidate.leg.options.approvePlan = () => false;
    let report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('failed');
    expect(f.candidate.actuate).not.toHaveBeenCalled();
    expect(report.measurements.syncReduction).toBeNull();
    const uncertain = pair();
    uncertain.candidate.setOutcome('uncertain');
    report = await runOfficeBenchmarkComparison(uncertain.input, uncertain.store);
    expect(report.status).toBe('failed');
    expect(report.candidate.effects[0]?.assessment).toBe('uncertain');
  });
  it('requires observed safety checks and scoped source/output equality rather than default zero defects', async () => {
    const f = pair();
    f.candidate.setChecks({ ...passedChecks, approvalEnforced: 'unavailable' });
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).not.toBe('passed');
    expect(report.checks.approvalEnforced).toBe('unavailable');
    const mismatch = pair();
    mismatch.candidate.setSource('different-input');
    const changed = await runOfficeBenchmarkComparison(mismatch.input, mismatch.store);
    expect(changed.checks.sameInput).toBe('failed');
    expect(changed.status).toBe('failed');
  });
  it('rejects probe scope changes and readback failure without exporting host error contents', async () => {
    const f = pair();
    f.candidate.leg.options.approvePlan = () => {
      f.candidate.setKey('different-document');
      return true;
    };
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('failed');
    expect(report.issueCodes).toContain('probe-changed-scope');
    const failed = pair();
    failed.candidate.failProbe();
    const broken = await runOfficeBenchmarkComparison(failed.input, failed.store);
    expect(broken.issueCodes).toContain('probe-failed');
    expect(JSON.stringify(broken)).not.toContain('private-host-error');
  });
  it('detects unrelated dispatch traces and refuses to attribute their window to the benchmark', async () => {
    const f = pair();
    f.candidate.contaminate();
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.candidate.syncMeasurementComplete).toBe(false);
    expect(report.issueCodes).toContain('trace-contamination');
    expect(report.measurements.syncReduction).toBeNull();
  });
  it('uses observed provider-body input bytes independently of command-query metrics', async () => {
    const f = pair();
    f.input.providerScope = 'replay';
    for (const leg of [f.baseline.leg, f.candidate.leg])
      leg.execution = {
        kind: 'workflow',
        workflow: {
          id: 'contract-review',
          scope: 'The selected clause',
          playbook: 'Use the requested test replacement.',
        },
      };
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('passed');
    expect(report.providerScope).toBe('replay');
    expect(report.baseline.modelCalls).toBe(1);
    expect(report.candidate.modelCalls).toBe(1);
    expect(report.baseline.inferenceInputBytes).toBe(f.baseline.meter.inputBytes);
    expect(report.baseline.inferenceInputComplete).toBe(true);
    expect(report.baseline.inferenceInputBytes).toBeGreaterThan(report.baseline.queryBytes);
  });
  it('leaves complete inference input unknown when a model run has no observed transport meter', async () => {
    const f = pair();
    f.input.providerScope = 'replay';
    for (const leg of [f.baseline.leg, f.candidate.leg]) {
      leg.execution = {
        kind: 'workflow',
        workflow: { id: 'contract-review', scope: 'Selected clause', playbook: 'Test' },
      };
      delete leg.inputMeter;
    }
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).toBe('passed');
    expect(report.baseline.inferenceInputBytes).toBeNull();
    expect(report.baseline.inferenceInputComplete).toBe(false);
    expect(report.measurements.inferenceInputReduction).toBeNull();
  });
  it('fails closed on cancellation and does not start an already cancelled native effect', async () => {
    const f = pair();
    const abort = new AbortController();
    abort.abort();
    f.input.signal = abort.signal;
    const report = await runOfficeBenchmarkComparison(f.input, f.store);
    expect(report.status).not.toBe('passed');
    expect(report.issueCodes).toContain('cancelled');
    expect(f.baseline.actuate).not.toHaveBeenCalled();
  });
});
