import { z } from 'zod';
import {
  capabilityModule,
  OfficeBenchmarkComparisonReportSchema,
  TenantWorkflowReportSchema,
  type OfficeBenchmarkComparisonReport,
} from '@ge/contracts';
import {
  NativeCertificationReportSchema,
  nativeFixtureExpectation,
} from './native-certification.js';
import { OfficeReleaseEvidenceSchema } from './office-readiness.js';

const commit = z.string().regex(/^[a-f0-9]{40}$/);
const surface = z.enum(['word', 'excel', 'powerpoint', 'outlook']);
const id = z.string().regex(/^[a-z][a-z0-9._-]{0,63}$/);
export const OFFICE_ACCEPTANCE_FIXTURES = {
  word: ['word-tracked-edit', 'word-table-readback'],
  excel: [
    'excel-grid-readback',
    'excel-table-readback',
    'excel-pivot-readback',
    'excel-filter-readback',
    'excel-validation-readback',
  ],
  powerpoint: ['powerpoint-slide-readback', 'powerpoint-layout-readback'],
  outlook: ['outlook-review-form', 'outlook-compose-readback'],
} as const;
const workflows = {
  word: 'contract-review',
  excel: 'workbook-close',
  powerpoint: 'board-pack-refresh',
  outlook: 'outlook-response',
} as const;
const modelRequired = (host: z.infer<typeof surface>): boolean =>
  host === 'word' || host === 'outlook';
const nativeChecks = [
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
] as const;
const localChecks = [
  'typecheck',
  'tests',
  'lint',
  'build',
  'generated',
  'secret-scan',
  'security-review',
] as const;
const ProfileSchema = OfficeReleaseEvidenceSchema.shape.profiles.element.extend({
  id,
  syncReductionTarget: z.number().min(0).max(1).default(0.5),
  inferenceInputReductionTarget: z.number().min(0).max(1).default(0.5),
});
export const OfficeAcceptancePlanSchema = z
  .object({
    version: z.literal(1),
    commit,
    requiredSurfaces: z.array(surface).min(1).max(4),
    profiles: z.array(ProfileSchema).max(64),
    minimumAttempts: z.number().int().min(20).max(1000).default(20),
  })
  .strict()
  .superRefine((plan, ctx) => {
    if (
      new Set(plan.requiredSurfaces).size !== plan.requiredSurfaces.length ||
      new Set(plan.profiles.map((profile) => profile.id)).size !== plan.profiles.length
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Surface and profile IDs must be unique.',
      });
    if (plan.profiles.length * plan.minimumAttempts > 4096)
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'The declared attempt matrix exceeds the 4,096-comparison evidence budget.',
      });
    for (const profile of plan.profiles) {
      if (
        !plan.requiredSurfaces.includes(profile.surface) ||
        profile.mode !== (profile.surface === 'outlook' ? 'compose' : 'write') ||
        /unknown|simulat|test-only|fake/i.test(profile.platform + ' ' + profile.build)
      )
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message: 'Advertise a matching native write/compose profile for each workflow.',
        });
    }
  });
export type OfficeAcceptancePlan = z.input<typeof OfficeAcceptancePlanSchema>;
export const OfficeAcceptanceEvidenceSchema = z
  .object({
    version: z.literal(1),
    commit,
    local: OfficeReleaseEvidenceSchema.shape.local,
    native: z.array(NativeCertificationReportSchema).max(1000),
    tenant: z.array(TenantWorkflowReportSchema).max(32),
    benchmarks: z.array(OfficeBenchmarkComparisonReportSchema).max(4096),
  })
  .strict()
  .superRefine((evidence, ctx) => {
    if (new TextEncoder().encode(JSON.stringify(evidence)).byteLength > 8 * 1024 * 1024)
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Acceptance evidence exceeds eight MiB.',
      });
    if (new Set(evidence.benchmarks.map((report) => report.id)).size !== evidence.benchmarks.length)
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'A benchmark comparison cannot be counted more than once.',
      });
  });
export type OfficeAcceptanceEvidence = z.input<typeof OfficeAcceptanceEvidenceSchema>;
export const OfficeAcceptanceBundleSchema = z
  .object({ plan: OfficeAcceptancePlanSchema, evidence: OfficeAcceptanceEvidenceSchema })
  .strict();
export type OfficeAcceptanceBundle = z.input<typeof OfficeAcceptanceBundleSchema>;
export interface OfficeAcceptanceCell {
  profileId: string;
  workflowId: (typeof workflows)[keyof typeof workflows];
  nativeFixtures: { passed: number; required: number };
  modelEvidence: 'passed' | 'missing' | 'failed' | 'not-required';
  attempts: number;
  successes: number;
  firstPassSuccessRate: number | null;
  syncReduction: number | null;
  inferenceInputReduction: number | null;
}
export interface OfficeAcceptanceAssessment {
  version: 1;
  commit: string;
  status: 'PASSED' | 'FAIL' | 'BLOCKED_EXTERNAL';
  failures: string[];
  missing: string[];
  cells: OfficeAcceptanceCell[];
  simulatedModelReports: number;
  excludedBenchmarks: number;
  evidenceOnly: true;
}
/** Empty evidence is a blocker, never a manufactured passing result or a guessed host profile. */
export function createOfficeAcceptanceBundle(
  head: string,
  requiredSurfaces: z.infer<typeof surface>[] = ['word', 'excel', 'powerpoint', 'outlook'],
): OfficeAcceptanceBundle {
  return OfficeAcceptanceBundleSchema.parse({
    plan: { version: 1, commit: head, requiredSurfaces, profiles: [], minimumAttempts: 20 },
    evidence: { version: 1, commit: head, local: [], native: [], tenant: [], benchmarks: [] },
  });
}
function matchesHost(
  profile: z.infer<typeof ProfileSchema>,
  host: OfficeBenchmarkComparisonReport['baseline']['host'],
): boolean {
  return Boolean(
    host &&
    host.surface === profile.surface &&
    host.platform === profile.platform &&
    host.build === profile.build &&
    host.mode === profile.mode,
  );
}
function reduction(baseline: number, candidate: number): number | null {
  return baseline > 0 ? 1 - candidate / baseline : null;
}
function equalReduction(left: number | null, right: number | null): boolean {
  return left === right || (left !== null && right !== null && Math.abs(left - right) < 1e-12);
}
function nativeOutcomeMatches(report: z.infer<typeof NativeCertificationReportSchema>): boolean {
  const expectation = nativeFixtureExpectation(report.fixtureId);
  return Boolean(
    expectation &&
    expectation.surface === report.surface &&
    expectation.expectedEffects.length === report.effects.length &&
    report.effects.every(
      (effect, index) => effect.ok && effect.kind === expectation.expectedEffects[index],
    ) &&
    (expectation.verifiedEffects ?? []).every(
      (index) => report.effects[index]?.verification === 'verified',
    ) &&
    (!expectation.reviewOpened || report.effects.at(-1)?.completion === 'review-opened'),
  );
}
/** Imported metadata is evidence, not an attestation. Every claimed pass is checked independently. */
export function assessOfficeAcceptance(
  raw: OfficeAcceptanceBundle,
  expectedCommit: string,
): OfficeAcceptanceAssessment {
  commit.parse(expectedCommit);
  const { plan, evidence } = OfficeAcceptanceBundleSchema.parse(raw);
  const failures: string[] = [];
  const missing: string[] = [];
  const cells: OfficeAcceptanceCell[] = [];
  if (plan.commit !== expectedCommit || evidence.commit !== expectedCommit)
    failures.push('The campaign is for a different commit.');
  for (const record of [
    ...evidence.local,
    ...evidence.native,
    ...evidence.tenant,
    ...evidence.benchmarks,
  ])
    if (record.commit !== expectedCommit) {
      failures.push('An imported report is for a different commit.');
      break;
    }
  for (const check of localChecks) {
    const records = evidence.local.filter(
      (record) => record.id === check && record.commit === expectedCommit,
    );
    if (records.some((record) => record.status === 'failed'))
      failures.push(`Local ${check} failed.`);
    if (!records.some((record) => record.status === 'passed'))
      missing.push(`Local ${check} at the exact commit.`);
  }
  for (const required of plan.requiredSurfaces)
    if (!plan.profiles.some((profile) => profile.surface === required))
      missing.push(`Advertised ${required} host/build profile.`);
  let excludedBenchmarks = 0;
  for (const report of evidence.benchmarks) {
    const profile = plan.profiles.find((profile) => profile.id === report.profileId);
    if (
      !profile ||
      report.hostScope !== 'native' ||
      report.providerScope === 'replay' ||
      (modelRequired(profile.surface) && report.providerScope !== 'live') ||
      !matchesHost(profile, report.baseline.host) ||
      !matchesHost(profile, report.candidate.host) ||
      report.workflowId !== workflows[profile.surface]
    )
      excludedBenchmarks++;
  }
  for (const profile of plan.profiles) {
    const cell: OfficeAcceptanceCell = {
      profileId: profile.id,
      workflowId: workflows[profile.surface],
      nativeFixtures: { passed: 0, required: OFFICE_ACCEPTANCE_FIXTURES[profile.surface].length },
      modelEvidence: modelRequired(profile.surface) ? 'missing' : 'not-required',
      attempts: 0,
      successes: 0,
      firstPassSuccessRate: null,
      syncReduction: null,
      inferenceInputReduction: null,
    };
    cells.push(cell);
    for (const fixtureId of OFFICE_ACCEPTANCE_FIXTURES[profile.surface]) {
      const reports = evidence.native.filter(
        (report) =>
          report.fixtureId === fixtureId &&
          report.surface === profile.surface &&
          report.commit === expectedCommit &&
          report.host?.platform === profile.platform &&
          report.host.build === profile.build &&
          report.host.mode === profile.mode,
      );
      if (
        reports.some(
          (report) =>
            report.status === 'failed' ||
            (report.status === 'passed' &&
              (!nativeChecks.every((check) => report.checks[check] === true) ||
                !Object.values(report.checks).every(Boolean) ||
                !nativeOutcomeMatches(report))),
        )
      )
        failures.push(`Native ${fixtureId} failed at ${profile.id}.`);
      if (
        reports.some(
          (report) =>
            report.status === 'passed' &&
            nativeChecks.every((check) => report.checks[check] === true) &&
            Object.values(report.checks).every(Boolean) &&
            nativeOutcomeMatches(report),
        )
      )
        cell.nativeFixtures.passed++;
      else missing.push(`Native ${fixtureId} at ${profile.id}.`);
    }
    if (modelRequired(profile.surface)) {
      for (const report of evidence.tenant.filter(
        (report) =>
          report.commit === expectedCommit && report.providerScope === 'live' && !report.dirty,
      )) {
        const attempts = report.attempts.filter(
          (attempt) =>
            attempt.workflowId === cell.workflowId &&
            attempt.surface === profile.surface &&
            attempt.modelRequired,
        );
        if (attempts.some((attempt) => attempt.status === 'failed')) cell.modelEvidence = 'failed';
        if (
          cell.modelEvidence !== 'failed' &&
          report.sourceFreshnessVerified &&
          report.approvalRefusalVerified &&
          attempts.some(
            (attempt) =>
              attempt.status === 'passed' &&
              attempt.modelCalls > 0 &&
              attempt.outcomeVerified &&
              attempt.requestRoutingVerified &&
              attempt.skillInvocationObserved,
          )
        )
          cell.modelEvidence = 'passed';
      }
      if (cell.modelEvidence === 'failed')
        failures.push(`Live model workflow failed for ${cell.workflowId}.`);
      else if (cell.modelEvidence !== 'passed')
        missing.push(`Observed live model workflow and skill route for ${cell.workflowId}.`);
    }
    const reports = evidence.benchmarks.filter(
      (report) =>
        report.commit === expectedCommit &&
        report.profileId === profile.id &&
        report.workflowId === cell.workflowId &&
        report.hostScope === 'native' &&
        report.providerScope !== 'replay' &&
        (!modelRequired(profile.surface) || report.providerScope === 'live') &&
        matchesHost(profile, report.baseline.host) &&
        matchesHost(profile, report.candidate.host),
    );
    let baseSyncs = 0;
    let nextSyncs = 0;
    let baseInput = 0;
    let nextInput = 0;
    let syncComplete = true;
    let inputComplete = true;
    for (const report of reports) {
      if (report.status !== 'unavailable') cell.attempts++;
      if (report.status === 'passed') cell.successes++;
      if (Object.values(report.checks).some((check) => check === 'failed'))
        failures.push(`Benchmark invariant failed at ${profile.id}.`);
      if (Object.values(report.checks).some((check) => check === 'unavailable'))
        missing.push(`Complete benchmark invariant observations at ${profile.id}.`);
      if (
        report.checks.artifactsMatch === 'passed' &&
        (!report.baseline.artifactHash ||
          report.baseline.artifactHash !== report.candidate.artifactHash)
      )
        failures.push(
          `Benchmark artifact hashes contradict their equality check at ${profile.id}.`,
        );
      if (
        report.checks.sameInput === 'passed' &&
        (!report.baseline.sourceHash ||
          report.baseline.sourceHash !== report.candidate.sourceHash ||
          !report.baseline.initialArtifactHash ||
          report.baseline.initialArtifactHash !== report.candidate.initialArtifactHash)
      )
        failures.push(`Benchmark inputs differ at ${profile.id}.`);
      if (
        report.baseline.probeId !== report.candidate.probeId ||
        report.baseline.evaluatorId !== report.candidate.evaluatorId
      )
        failures.push(`Benchmark probe/evaluator differs at ${profile.id}.`);
      if (
        report.status === 'passed' &&
        (report.issueCodes.length > 0 ||
          report.baseline.issueCodes.length > 0 ||
          report.candidate.issueCodes.length > 0 ||
          !report.baseline.scopeHash ||
          !report.candidate.scopeHash ||
          !report.baseline.checks ||
          !report.candidate.checks ||
          Object.values(report.baseline.checks).some((check) => check !== 'passed') ||
          Object.values(report.candidate.checks).some((check) => check !== 'passed') ||
          !report.baseline.effects.length ||
          !report.candidate.effects.length ||
          report.baseline.status !== 'passed' ||
          report.candidate.status !== 'passed' ||
          report.baseline.effects.some(
            (effect) =>
              !['verified', 'review-opened'].includes(effect.assessment) ||
              !capabilityModule(profile.surface, effect.kind) ||
              (effect.assessment === 'review-opened' &&
                (profile.surface !== 'outlook' ||
                  !['create-mail', 'reply-mail'].includes(effect.kind))),
          ) ||
          report.candidate.effects.some(
            (effect) =>
              !['verified', 'review-opened'].includes(effect.assessment) ||
              !capabilityModule(profile.surface, effect.kind) ||
              (effect.assessment === 'review-opened' &&
                (profile.surface !== 'outlook' ||
                  !['create-mail', 'reply-mail'].includes(effect.kind))),
          ))
      )
        failures.push(`A benchmark pass contains an incomplete native outcome at ${profile.id}.`);
      for (const check of [
        'targetIntegrity',
        'approvalEnforced',
        'completionTruth',
        'freshnessRefusal',
      ] as const)
        if (
          report.checks[check] === 'passed' &&
          (report.baseline.checks?.[check] !== 'passed' ||
            report.candidate.checks?.[check] !== 'passed')
        )
          failures.push(`Benchmark leg checks contradict the comparison at ${profile.id}.`);
      if (
        modelRequired(profile.surface) &&
        (!report.baseline.modelCalls || !report.candidate.modelCalls)
      )
        missing.push(`Native workflow model execution at ${profile.id}.`);
      const completeSync =
        report.baseline.syncMeasurementComplete && report.candidate.syncMeasurementComplete;
      syncComplete &&= completeSync;
      baseSyncs += report.baseline.nativeSyncs;
      nextSyncs += report.candidate.nativeSyncs;
      if (
        !equalReduction(
          report.measurements.syncReduction,
          completeSync
            ? reduction(report.baseline.nativeSyncs, report.candidate.nativeSyncs)
            : null,
        )
      )
        failures.push(`Benchmark sync reduction is inconsistent at ${profile.id}.`);
      const completeInput =
        report.baseline.inferenceInputComplete &&
        report.candidate.inferenceInputComplete &&
        report.baseline.inferenceInputBytes !== null &&
        report.candidate.inferenceInputBytes !== null;
      inputComplete &&= completeInput;
      if (completeInput) {
        baseInput += report.baseline.inferenceInputBytes!;
        nextInput += report.candidate.inferenceInputBytes!;
      }
      if (
        !equalReduction(
          report.measurements.inferenceInputReduction,
          completeInput
            ? reduction(report.baseline.inferenceInputBytes!, report.candidate.inferenceInputBytes!)
            : null,
        )
      )
        failures.push(`Benchmark inference reduction is inconsistent at ${profile.id}.`);
    }
    if (cell.attempts < plan.minimumAttempts)
      missing.push(`At least ${plan.minimumAttempts} native comparison attempts at ${profile.id}.`);
    cell.firstPassSuccessRate = cell.attempts ? cell.successes / cell.attempts : null;
    if (cell.firstPassSuccessRate !== null && cell.firstPassSuccessRate < 0.95)
      failures.push(`First-pass success below 95% at ${profile.id}.`);
    if (reports.length && syncComplete) {
      cell.syncReduction = reduction(baseSyncs, nextSyncs);
      if (
        (cell.syncReduction !== null && cell.syncReduction + 1e-12 < profile.syncReductionTarget) ||
        (!baseSyncs && nextSyncs > 0)
      )
        failures.push(`Native sync target missed at ${profile.id}.`);
    } else if (reports.length) missing.push(`Complete native sync measurements at ${profile.id}.`);
    if (reports.length && inputComplete) {
      cell.inferenceInputReduction = reduction(baseInput, nextInput);
      if (
        (cell.inferenceInputReduction !== null &&
          cell.inferenceInputReduction + 1e-12 < profile.inferenceInputReductionTarget) ||
        (!baseInput && nextInput > 0)
      )
        failures.push(`Inference input target missed at ${profile.id}.`);
    } else if (reports.length)
      missing.push(`Complete provider input measurements at ${profile.id}.`);
  }
  const uniqueFailures = [...new Set(failures)];
  const uniqueMissing = [...new Set(missing)];
  return {
    version: 1,
    commit: expectedCommit,
    status: uniqueFailures.length ? 'FAIL' : uniqueMissing.length ? 'BLOCKED_EXTERNAL' : 'PASSED',
    failures: uniqueFailures,
    missing: uniqueMissing,
    cells,
    simulatedModelReports: evidence.tenant.length,
    excludedBenchmarks,
    evidenceOnly: true,
  };
}

/** Normalize independently exported reports; no conversion of simulator reports into native evidence. */
export function collectOfficeAcceptanceEvidence(
  planRaw: OfficeAcceptancePlan,
  documents: readonly unknown[],
): OfficeAcceptanceBundle {
  const plan = OfficeAcceptancePlanSchema.parse(planRaw);
  const evidence: OfficeAcceptanceEvidence = {
    version: 1,
    commit: plan.commit,
    local: [],
    native: [],
    tenant: [],
    benchmarks: [],
  };
  const add = (document: unknown): void => {
    const native = NativeCertificationReportSchema.safeParse(document);
    if (native.success) {
      evidence.native.push(native.data);
      return;
    }
    const tenant = TenantWorkflowReportSchema.safeParse(document);
    if (tenant.success) {
      evidence.tenant.push(tenant.data);
      return;
    }
    const benchmark = OfficeBenchmarkComparisonReportSchema.safeParse(document);
    if (benchmark.success) {
      evidence.benchmarks.push(benchmark.data);
      return;
    }
    const local = z
      .object({ version: z.literal(1), commit, local: OfficeReleaseEvidenceSchema.shape.local })
      .strict()
      .safeParse(document);
    if (local.success) {
      if (local.data.commit !== plan.commit)
        throw new Error('Local checks belong to a different commit.');
      evidence.local.push(...local.data.local);
      return;
    }
    throw new Error(
      'Unsupported or invalid acceptance report. Source documents and credentials are not accepted.',
    );
  };
  if (documents.length > 1000) throw new Error('Import at most 1,000 acceptance report files.');
  for (const document of documents) {
    if (Array.isArray(document)) {
      if (document.length > 1000) throw new Error('A native report array is too large.');
      document.forEach(add);
    } else add(document);
  }
  return OfficeAcceptanceBundleSchema.parse({ plan, evidence });
}
