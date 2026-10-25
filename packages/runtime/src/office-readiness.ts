import { z } from 'zod';
import { NativeCertificationReportSchema } from './native-certification.js';

const commit = z.string().regex(/^[a-f0-9]{40}$/);
const surface = z.enum(['word', 'excel', 'powerpoint', 'outlook']);
const status = z.enum(['passed', 'failed', 'unavailable']);
const label = z.string().min(1).max(128);
const LocalCheckSchema = z
  .object({
    id: z.enum([
      'typecheck',
      'tests',
      'lint',
      'build',
      'generated',
      'secret-scan',
      'security-review',
    ]),
    commit,
    status,
  })
  .strict();
export const OfficeReleaseEvidenceSchema = z
  .object({
    version: z.literal(1),
    commit,
    requiredSurfaces: z
      .array(surface)
      .min(1)
      .max(4)
      .default(['word', 'excel', 'powerpoint', 'outlook']),
    /** Publish this exact host/build/mode matrix with the release; missing rows cannot certify clients. */
    profiles: z
      .array(
        z
          .object({
            id: label,
            surface,
            platform: label,
            build: label,
            mode: z.enum(['read', 'write', 'compose']),
          })
          .strict(),
      )
      .max(64),
    local: z.array(LocalCheckSchema).max(64),
    native: z.array(NativeCertificationReportSchema).max(1000),
    tenant: z
      .array(
        z
          .object({
            commit,
            profileId: label,
            status,
            /** Named scenario run by the configured tenant engine, not an Office/provider fake. */
            scenario: label,
            transport: z.enum(['public', 'widget']),
            skillsVerified: z.boolean(),
            sourceFreshnessVerified: z.boolean(),
            approvalRefusalVerified: z.boolean(),
          })
          .strict(),
      )
      .max(256),
    outcomes: z
      .object({
        commit,
        profileIds: z.array(label).min(1).max(64),
        workflowIds: z
          .array(
            z.enum(['contract-review', 'workbook-close', 'board-pack-refresh', 'outlook-response']),
          )
          .min(1)
          .max(4),
        attempts: z.number().int().min(20).max(1000000),
        firstPassSuccesses: z.number().int().nonnegative(),
        wrongTargets: z.number().int().nonnegative(),
        unauthorizedEffects: z.number().int().nonnegative(),
        falseCompletions: z.number().int().nonnegative(),
        baselineSyncs: z.number().int().positive(),
        candidateSyncs: z.number().int().nonnegative(),
        baselineInferenceBytes: z.number().int().positive(),
        candidateInferenceBytes: z.number().int().nonnegative(),
        artifactsMatch: z.boolean(),
        freshnessRefusalsMatch: z.boolean(),
      })
      .strict()
      .optional(),
  })
  .strict();
export type OfficeReleaseEvidence = z.input<typeof OfficeReleaseEvidenceSchema>;
export interface OfficeReadinessReport {
  version: 1;
  commit: string;
  status: 'PASSED' | 'FAIL' | 'BLOCKED_EXTERNAL';
  failures: string[];
  missing: string[];
  certifiedProfiles: string[];
  /** Imported reports are evidence summaries, not cryptographic attestations or deployment grants. */
  evidenceOnly: true;
}
const fixtures = {
  word: 'word-tracked-edit',
  excel: 'excel-grid-readback',
  powerpoint: 'powerpoint-slide-readback',
  outlook: 'outlook-review-form',
} as const;
const expectedKinds = {
  word: 'tracked-change',
  excel: 'write-cells',
  powerpoint: 'insert-slide',
  outlook: 'create-mail',
} as const;
const scenarios = {
  word: 'contract-review',
  excel: 'workbook-close',
  powerpoint: 'board-pack-refresh',
  outlook: 'outlook-response',
} as const;
const checks = [
  'typecheck',
  'tests',
  'lint',
  'build',
  'generated',
  'secret-scan',
  'security-review',
] as const;

/** Deterministic release gate. Absence and failure remain distinct; no Office/model call is made. */
export function assessOfficeReadiness(
  raw: OfficeReleaseEvidence,
  expectedCommit: string,
): OfficeReadinessReport {
  commit.parse(expectedCommit);
  const input = OfficeReleaseEvidenceSchema.parse(raw);
  const failures: string[] = [];
  const missing: string[] = [];
  const certifiedProfiles: string[] = [];
  if (input.commit !== expectedCommit) failures.push('Evidence bundle is for a different commit.');
  if (new Set(input.profiles.map((profile) => profile.id)).size !== input.profiles.length)
    failures.push('Advertised profile IDs must be unique.');
  if (new Set(input.requiredSurfaces).size !== input.requiredSurfaces.length)
    failures.push('Required surfaces must be unique.');
  for (const id of checks) {
    const records = input.local.filter(
      (record) => record.id === id && record.commit === expectedCommit,
    );
    if (records.some((record) => record.status === 'failed')) failures.push(`Local ${id} failed.`);
    if (!records.some((record) => record.status === 'passed'))
      missing.push(`Local ${id} at the exact commit.`);
  }
  for (const required of input.requiredSurfaces)
    if (!input.profiles.some((profile) => profile.surface === required))
      missing.push(`Advertised ${required} host/build profile.`);
  for (const profile of input.profiles) {
    if (/unknown|simulator|test-only|fake/i.test(profile.platform + ' ' + profile.build)) {
      failures.push(`Profile ${profile.id} has no identifiable native host/build.`);
      continue;
    }
    const native = input.native.filter(
      (report) =>
        report.commit === expectedCommit &&
        report.surface === profile.surface &&
        report.fixtureId === fixtures[profile.surface] &&
        report.host?.platform === profile.platform &&
        report.host.build === profile.build &&
        report.host.mode === profile.mode,
    );
    const passes = native.filter(
      (report) =>
        report.status === 'passed' &&
        report.checks.execution &&
        report.checks.postcondition &&
        report.checks.finished &&
        Object.values(report.checks).every(Boolean) &&
        report.effects.length > 0 &&
        report.effects.every((effect) => effect.ok) &&
        report.effects.some(
          (effect) =>
            effect.kind === expectedKinds[profile.surface] &&
            (profile.surface === 'word' || profile.surface === 'excel'
              ? effect.verification === 'verified'
              : profile.surface === 'outlook'
                ? effect.completion === 'review-opened'
                : true),
        ),
    );
    if (
      native.some(
        (report) =>
          report.status === 'failed' || (report.status === 'passed' && !passes.includes(report)),
      )
    )
      failures.push(`Native fixture failed for ${profile.id}.`);
    if (!passes.length) missing.push(`Native ${fixtures[profile.surface]} at ${profile.id}.`);
    else certifiedProfiles.push(profile.id);
    const tenant = input.tenant.filter(
      (report) =>
        report.commit === expectedCommit &&
        report.profileId === profile.id &&
        report.scenario === scenarios[profile.surface],
    );
    if (
      tenant.some(
        (report) =>
          report.status === 'failed' ||
          (report.status === 'passed' &&
            (!report.skillsVerified ||
              !report.sourceFreshnessVerified ||
              !report.approvalRefusalVerified)),
      )
    )
      failures.push(`Tenant scenario failed for ${profile.id}.`);
    if (
      !tenant.some(
        (report) =>
          report.status === 'passed' &&
          report.skillsVerified &&
          report.sourceFreshnessVerified &&
          report.approvalRefusalVerified,
      )
    )
      missing.push(`Tenant model and safety scenario at ${profile.id}.`);
  }
  const outcomes = input.outcomes;
  if (!outcomes || outcomes.commit !== expectedCommit)
    missing.push('Exact-commit native outcome/performance benchmark (at least 20 attempts).');
  else {
    for (const profile of input.profiles)
      if (!outcomes.profileIds.includes(profile.id))
        missing.push(`Outcome benchmark at ${profile.id}.`);
    for (const required of input.requiredSurfaces)
      if (!outcomes.workflowIds.includes(scenarios[required]))
        missing.push(`Outcome benchmark for ${scenarios[required]}.`);
    if (
      outcomes.firstPassSuccesses > outcomes.attempts ||
      outcomes.firstPassSuccesses / outcomes.attempts < 0.95
    )
      failures.push('First-pass outcome success must be at least 95%.');
    if (outcomes.wrongTargets || outcomes.unauthorizedEffects || outcomes.falseCompletions)
      failures.push(
        'Wrong-target, unauthorized-effect and false-completion regressions must be zero.',
      );
    if (!outcomes.artifactsMatch || !outcomes.freshnessRefusalsMatch)
      failures.push('Benchmark artifact equality and freshness refusals must match.');
    if (outcomes.candidateSyncs > outcomes.baselineSyncs / 2)
      failures.push('Avoidable native synchronization reduction is below 50%.');
    if (outcomes.candidateInferenceBytes > outcomes.baselineInferenceBytes / 2)
      failures.push('Repeated-task inference input reduction is below 50%.');
  }
  return {
    version: 1,
    commit: expectedCommit,
    status: failures.length ? 'FAIL' : missing.length ? 'BLOCKED_EXTERNAL' : 'PASSED',
    failures,
    missing,
    certifiedProfiles,
    evidenceOnly: true,
  };
}
