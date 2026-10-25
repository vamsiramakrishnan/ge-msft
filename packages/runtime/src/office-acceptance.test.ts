import { describe, expect, it } from 'vitest';
import type { OfficeBenchmarkComparisonReport, TenantWorkflowReport } from '@ge/contracts';
import {
  nativeFixtureExpectation,
  type NativeCertificationReport,
} from './native-certification.js';
import {
  assessOfficeAcceptance,
  collectOfficeAcceptanceEvidence,
  createOfficeAcceptanceBundle,
  OFFICE_ACCEPTANCE_FIXTURES,
  type OfficeAcceptanceBundle,
} from './office-acceptance.js';

const commit = 'a'.repeat(40);
const hash = (letter: string) => 'sha256:' + letter.repeat(64);
function tenant(): TenantWorkflowReport {
  const workflows = [
    'contract-review',
    'workbook-close',
    'board-pack-refresh',
    'outlook-response',
  ] as const;
  const surfaces = ['word', 'excel', 'powerpoint', 'outlook'] as const;
  return {
    version: 1,
    commit,
    dirty: false,
    providerScope: 'live',
    hostScope: 'simulated',
    identityScope: 'explicit-developer-token',
    transport: 'public',
    status: 'PASSED',
    sourceFreshnessVerified: true,
    approvalRefusalVerified: true,
    attempts: workflows.map((workflowId, index) => ({
      workflowId,
      surface: surfaces[index]!,
      repetition: 1,
      status: 'passed',
      modelRequired: index === 0 || index === 3,
      modelCalls: index === 0 || index === 3 ? 1 : 0,
      requestRoutingVerified: index === 0 || index === 3,
      skillInvocationObserved: index === 0 || index === 3,
      outcomeVerified: true,
      durationMs: 10,
      issueCodes: [],
    })),
    summary: {
      requested: 4,
      passed: 4,
      failed: 0,
      unavailable: 0,
      modelCalls: 2,
      durationMs: 40,
      p50Ms: 10,
      p95Ms: 10,
    },
  };
}
function native(fixtureId: string): NativeCertificationReport {
  const expectation = nativeFixtureExpectation(fixtureId)!;
  return {
    version: 1,
    fixtureVersion: 1,
    fixtureId,
    commit,
    surface: expectation.surface,
    host: {
      platform: 'PC',
      build: '16.0.18000',
      mode: 'write',
      requirements: { 'WordApi 1.4': true },
    },
    startedAt: '2026-10-25T00:00:00Z',
    durationMs: 10,
    status: 'passed',
    nativeSyncs: 2,
    checks: Object.fromEntries(
      [
        'nativeHost',
        'documentIdentity',
        'mode',
        'support',
        'finished',
        'approved',
        'correlated',
        'expectedEffects',
        'scopeStable',
        'hostStable',
        'execution',
        'expectedEffect',
        'postcondition',
      ].map((check) => [check, true]),
    ),
    effects: expectation.expectedEffects.map((kind, index) => ({
      kind,
      ok: true,
      ...(expectation.verifiedEffects?.includes(index) ? { verification: 'verified' } : {}),
      ...(expectation.reviewOpened ? { completion: 'review-opened' } : {}),
    })),
  };
}
function benchmark(): OfficeBenchmarkComparisonReport {
  const checks = {
    targetIntegrity: 'passed',
    approvalEnforced: 'passed',
    completionTruth: 'passed',
    freshnessRefusal: 'passed',
  } as const;
  const leg = {
    status: 'passed' as const,
    host: {
      surface: 'word' as const,
      platform: 'PC',
      build: '16.0.18000',
      mode: 'write' as const,
      requirements: {},
    },
    durationMs: 10,
    nativeSyncs: 4,
    syncMeasurementComplete: true,
    modelCalls: 1,
    queryBytes: 800,
    queryBytesComplete: true,
    inferenceInputBytes: 1000,
    inferenceInputComplete: true,
    effects: [{ kind: 'tracked-change' as const, assessment: 'verified' as const }],
    scopeHash: hash('a'),
    sourceHash: hash('b'),
    initialArtifactHash: hash('c'),
    artifactHash: hash('d'),
    probeId: 'word-output-v1',
    evaluatorId: 'word-policy-v1',
    checks,
    issueCodes: [],
  };
  return {
    version: 1,
    id: crypto.randomUUID(),
    commit,
    workflowId: 'contract-review',
    profileId: 'word-pc',
    hostScope: 'native',
    providerScope: 'live',
    status: 'passed',
    baseline: leg,
    candidate: {
      ...structuredClone(leg),
      nativeSyncs: 2,
      queryBytes: 400,
      inferenceInputBytes: 500,
    },
    checks: { ...checks, sameInput: 'passed', artifactsMatch: 'passed', sourceFreshness: 'passed' },
    measurements: { syncReduction: 0.5, inferenceInputReduction: 0.5 },
    issueCodes: [],
  };
}
function complete(): OfficeAcceptanceBundle {
  const bundle = createOfficeAcceptanceBundle(commit, ['word']);
  bundle.plan.profiles.push({
    id: 'word-pc',
    surface: 'word',
    platform: 'PC',
    build: '16.0.18000',
    mode: 'write',
    syncReductionTarget: 0.5,
    inferenceInputReductionTarget: 0.5,
  });
  bundle.evidence.local = [
    'typecheck',
    'tests',
    'lint',
    'build',
    'generated',
    'secret-scan',
    'security-review',
  ].map((id) => ({
    id: id as OfficeAcceptanceBundle['evidence']['local'][number]['id'],
    commit,
    status: 'passed',
  }));
  bundle.evidence.native = OFFICE_ACCEPTANCE_FIXTURES.word.map(native);
  bundle.evidence.tenant = [tenant()];
  bundle.evidence.benchmarks = Array.from({ length: 20 }, benchmark);
  return bundle;
}
describe('Office acceptance campaign (synthetic imported metadata, no host execution)', () => {
  it('starts unavailable and never invents a host, run or passing local check', () => {
    const result = assessOfficeAcceptance(createOfficeAcceptanceBundle(commit), commit);
    expect(result).toMatchObject({
      status: 'BLOCKED_EXTERNAL',
      failures: [],
      cells: [],
      excludedBenchmarks: 0,
      evidenceOnly: true,
    });
    expect(result.missing).toContain('Advertised word host/build profile.');
  });
  it('requires the exact native fixture and measured workflow matrix for each advertised profile', () => {
    const result = assessOfficeAcceptance(complete(), commit);
    expect(result).toMatchObject({
      status: 'PASSED',
      failures: [],
      missing: [],
      cells: [
        {
          profileId: 'word-pc',
          nativeFixtures: { passed: 2, required: 2 },
          modelEvidence: 'passed',
          attempts: 20,
          successes: 20,
          firstPassSuccessRate: 1,
          syncReduction: 0.5,
          inferenceInputReduction: 0.5,
        },
      ],
    });
  });
  it('does not promote replay or live-model/simulated-host checks into native comparisons', () => {
    const bundle = complete();
    bundle.evidence.benchmarks.forEach((report) => {
      report.hostScope = 'simulated';
    });
    const result = assessOfficeAcceptance(bundle, commit);
    expect(result.status).toBe('BLOCKED_EXTERNAL');
    expect(result.excludedBenchmarks).toBe(20);
    expect(result.cells[0]!.attempts).toBe(0);
    bundle.evidence.benchmarks.forEach((report) => {
      report.hostScope = 'native';
      report.providerScope = 'replay';
    });
    expect(assessOfficeAcceptance(bundle, commit).cells[0]!.attempts).toBe(0);
  });
  it('cannot reuse one successful profile to certify another advertised build', () => {
    const bundle = complete();
    bundle.plan.profiles.push({
      id: 'word-mac',
      surface: 'word',
      platform: 'Mac',
      build: '16.0.19000',
      mode: 'write',
    });
    const result = assessOfficeAcceptance(bundle, commit);
    expect(result.status).toBe('BLOCKED_EXTERNAL');
    expect(result.cells[1]).toMatchObject({
      attempts: 0,
      nativeFixtures: { passed: 0, required: 2 },
    });
  });
  it('rejects duplicate comparison IDs instead of counting copied evidence repeatedly', () => {
    const bundle = complete();
    bundle.evidence.benchmarks[1]!.id = bundle.evidence.benchmarks[0]!.id;
    expect(() => assessOfficeAcceptance(bundle, commit)).toThrow('more than once');
  });
  it('checks actual receipt sequences and verification even when every imported check says true', () => {
    const bundle = complete();
    bundle.evidence.native[0]!.effects.push({ kind: 'set-body', ok: true });
    const result = assessOfficeAcceptance(bundle, commit);
    expect(result.status).toBe('FAIL');
    expect(result.cells[0]!.nativeFixtures.passed).toBe(1);
    bundle.evidence.native[0] = native('word-tracked-edit');
    bundle.evidence.native[0]!.effects[1]!.verification = 'unknown';
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
  });
  it('will not let an older passing report or a newer pass conceal a current failed case', () => {
    const bundle = complete();
    bundle.evidence.native.push({ ...native('word-tracked-edit'), status: 'failed' });
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
    bundle.evidence.native.pop();
    bundle.evidence.native[0]!.commit = 'b'.repeat(40);
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
  });
  it('requires observed routing and skill use, not configured mounting or replay success', () => {
    const bundle = complete();
    bundle.evidence.tenant[0]!.providerScope = 'replay';
    bundle.evidence.tenant[0]!.identityScope = 'synthetic-replay';
    expect(assessOfficeAcceptance(bundle, commit).cells[0]!.modelEvidence).toBe('missing');
    const noObservation = tenant();
    noObservation.status = 'BLOCKED_EXTERNAL';
    noObservation.attempts[0]!.status = 'unavailable';
    noObservation.attempts[0]!.skillInvocationObserved = false;
    noObservation.attempts[0]!.issueCodes = ['skill_invocation_unobserved'];
    noObservation.summary.passed = 3;
    noObservation.summary.unavailable = 1;
    bundle.evidence.tenant = [noObservation];
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('BLOCKED_EXTERNAL');
  });
  it('fails safety/equality contradictions independently of headline outcome and performance', () => {
    const bundle = complete();
    bundle.evidence.benchmarks[0]!.checks.approvalEnforced = 'failed';
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
    bundle.evidence.benchmarks[0] = benchmark();
    bundle.evidence.benchmarks[0]!.candidate.artifactHash = hash('e');
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
  });
  it('does not substitute command-query bytes for measured provider input or accept invented reductions', () => {
    const bundle = complete();
    bundle.evidence.benchmarks[0]!.baseline.inferenceInputComplete = false;
    bundle.evidence.benchmarks[0]!.baseline.inferenceInputBytes = null;
    bundle.evidence.benchmarks[0]!.measurements.inferenceInputReduction = null;
    const result = assessOfficeAcceptance(bundle, commit);
    expect(result.status).toBe('BLOCKED_EXTERNAL');
    expect(result.cells[0]!.inferenceInputReduction).toBeNull();
    bundle.evidence.benchmarks[0] = benchmark();
    bundle.evidence.benchmarks[0]!.measurements.syncReduction = 0.9;
    expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
  });
  it('imports only typed reports and preserves every execution origin', () => {
    const plan = complete().plan;
    const bundle = collectOfficeAcceptanceEvidence(plan, [
      [native('word-tracked-edit')],
      tenant(),
      benchmark(),
    ]);
    expect(bundle.evidence.tenant[0]!.hostScope).toBe('simulated');
    expect(bundle.evidence.benchmarks).toHaveLength(1);
    expect(() => collectOfficeAcceptanceEvidence(plan, [{ token: 'secret', version: 1 }])).toThrow(
      'Unsupported or invalid',
    );
    expect(() =>
      collectOfficeAcceptanceEvidence(plan, [{ ...tenant(), extraDocumentText: 'private' }]),
    ).toThrow('Unsupported or invalid');
  });
  it('rejects empty receipts, missing probe scope and contamination behind a claimed pass', () => {
    for (const mutate of [
      (report: OfficeBenchmarkComparisonReport) => {
        report.candidate.effects = [];
      },
      (report: OfficeBenchmarkComparisonReport) => {
        delete report.baseline.scopeHash;
      },
      (report: OfficeBenchmarkComparisonReport) => {
        delete report.candidate.checks;
      },
      (report: OfficeBenchmarkComparisonReport) => {
        report.issueCodes = ['trace-contamination'];
      },
      (report: OfficeBenchmarkComparisonReport) => {
        report.baseline.issueCodes = ['probe-failed'];
      },
      (report: OfficeBenchmarkComparisonReport) => {
        report.candidate.effects[0]!.kind = 'set-body';
      },
    ]) {
      const bundle = complete();
      mutate(bundle.evidence.benchmarks[0]!);
      expect(assessOfficeAcceptance(bundle, commit).status).toBe('FAIL');
    }
  });
});
