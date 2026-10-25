import { z } from 'zod';
import { ActuationKindSchema } from './capability.js';
import { HostCapabilityProfileSchema } from './host-profile.js';

export const OfficeBenchmarkCheckSchema = z.enum(['passed', 'failed', 'unavailable']);
export type OfficeBenchmarkCheck = z.infer<typeof OfficeBenchmarkCheckSchema>;
export const OfficeBenchmarkWorkflowSchema = z.enum([
  'contract-review',
  'workbook-close',
  'board-pack-refresh',
  'outlook-response',
]);
const hash = z.string().regex(/^sha256:[a-f0-9]{64}$/);
const publicId = z.string().regex(/^[a-z][a-z0-9._-]{0,63}$/);
export const OfficeBenchmarkIssueSchema = z.enum([
  'unavailable-host',
  'wrong-surface',
  'invalid-origin',
  'execution-failed',
  'cancelled',
  'incomplete',
  'receipt-mismatch',
  'unverified-effect',
  'probe-failed',
  'probe-changed-scope',
  'source-changed',
  'different-input',
  'different-artifact',
  'target-check',
  'approval-check',
  'completion-check',
  'freshness-check',
  'trace-contamination',
  'trace-limit',
  'missing-metrics',
  'missing-approval',
  'task-overlap',
  'host-changed',
]);
export const OfficeBenchmarkObservedChecksSchema = z
  .object({
    targetIntegrity: OfficeBenchmarkCheckSchema,
    approvalEnforced: OfficeBenchmarkCheckSchema,
    completionTruth: OfficeBenchmarkCheckSchema,
    freshnessRefusal: OfficeBenchmarkCheckSchema,
  })
  .strict();
export const OfficeBenchmarkLegSchema = z
  .object({
    status: z.enum(['passed', 'failed', 'unavailable']),
    host: HostCapabilityProfileSchema.optional(),
    durationMs: z.number().finite().nonnegative(),
    nativeSyncs: z.number().int().nonnegative(),
    /** Counts are observed in the caller-confirmed exclusive process window, not inferred from a callback. */
    syncMeasurementComplete: z.boolean(),
    modelCalls: z.number().int().nonnegative(),
    queryBytes: z.number().int().nonnegative(),
    queryBytesComplete: z.boolean(),
    /** Complete serialized provider request-body bytes when an observed transport meter is supplied. */
    inferenceInputBytes: z.number().int().nonnegative().nullable(),
    inferenceInputComplete: z.boolean(),
    effects: z
      .array(
        z
          .object({
            kind: ActuationKindSchema,
            assessment: z.enum([
              'verified',
              'review-opened',
              'unverified',
              'uncertain',
              'rejected',
              'failed',
            ]),
          })
          .strict(),
      )
      .max(64),
    scopeHash: hash.optional(),
    sourceHash: hash.optional(),
    initialArtifactHash: hash.optional(),
    artifactHash: hash.optional(),
    probeId: publicId,
    evaluatorId: publicId,
    checks: OfficeBenchmarkObservedChecksSchema.optional(),
    issueCodes: z.array(OfficeBenchmarkIssueSchema).max(32),
  })
  .strict();
export type OfficeBenchmarkLeg = z.infer<typeof OfficeBenchmarkLegSchema>;
/** Metadata only. Reports never contain prompts, source values, target IDs, receipt bodies or bearer credentials. */
export const OfficeBenchmarkComparisonReportSchema = z
  .object({
    version: z.literal(1),
    id: z.string().uuid(),
    commit: z.string().regex(/^[a-f0-9]{40}$/),
    workflowId: OfficeBenchmarkWorkflowSchema,
    profileId: publicId,
    hostScope: z.enum(['native', 'simulated']),
    /** A declaration from trusted provider wiring, never inferred from mounted skill configuration. */
    providerScope: z.enum(['live', 'replay', 'none']),
    status: z.enum(['passed', 'failed', 'unavailable']),
    baseline: OfficeBenchmarkLegSchema,
    candidate: OfficeBenchmarkLegSchema,
    checks: OfficeBenchmarkObservedChecksSchema.extend({
      sameInput: OfficeBenchmarkCheckSchema,
      artifactsMatch: OfficeBenchmarkCheckSchema,
      sourceFreshness: OfficeBenchmarkCheckSchema,
    }).strict(),
    measurements: z
      .object({
        /** Null means no reduction claim (missing/incomplete metrics or a zero baseline). */
        syncReduction: z.number().finite().nullable(),
        inferenceInputReduction: z.number().finite().nullable(),
      })
      .strict(),
    issueCodes: z.array(OfficeBenchmarkIssueSchema).max(32),
  })
  .strict();
export type OfficeBenchmarkComparisonReport = z.infer<typeof OfficeBenchmarkComparisonReportSchema>;
