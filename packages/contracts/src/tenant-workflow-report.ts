import { z } from 'zod';

export const TenantWorkflowIdSchema = z.enum([
  'contract-review',
  'workbook-close',
  'board-pack-refresh',
  'outlook-response',
]);
export const TenantWorkflowIssueSchema = z.enum([
  'configuration_missing',
  'provider_unavailable',
  'query_budget',
  'routing_unobserved',
  'skill_invocation_unobserved',
  'outcome_mismatch',
  'approval_bypass',
  'stale_source_bypass',
  'unexpected_error',
  'dirty_worktree',
  'cancelled',
]);
export const TenantWorkflowAttemptSchema = z
  .object({
    workflowId: TenantWorkflowIdSchema,
    surface: z.enum(['word', 'excel', 'powerpoint', 'outlook']),
    repetition: z.number().int().min(1).max(5),
    status: z.enum(['passed', 'failed', 'unavailable']),
    modelRequired: z.boolean(),
    modelCalls: z.number().int().nonnegative().max(24),
    requestRoutingVerified: z.boolean(),
    skillInvocationObserved: z.boolean(),
    outcomeVerified: z.boolean(),
    durationMs: z.number().finite().nonnegative(),
    issueCodes: z.array(TenantWorkflowIssueSchema).max(16),
  })
  .strict();
export const TenantWorkflowReportSchema = z
  .object({
    version: z.literal(1),
    commit: z.string().regex(/^[a-f0-9]{40}$/),
    dirty: z.boolean(),
    providerScope: z.enum(['live', 'replay', 'unavailable']),
    transport: z.literal('public'),
    hostScope: z.literal('simulated'),
    identityScope: z.enum(['explicit-developer-token', 'synthetic-replay', 'unavailable']),
    status: z.enum(['PASSED', 'FAIL', 'BLOCKED_EXTERNAL']),
    sourceFreshnessVerified: z.boolean(),
    approvalRefusalVerified: z.boolean(),
    attempts: z.array(TenantWorkflowAttemptSchema).max(20),
    summary: z
      .object({
        requested: z.number().int().min(4).max(20),
        passed: z.number().int().nonnegative().max(20),
        failed: z.number().int().nonnegative().max(20),
        unavailable: z.number().int().nonnegative().max(20),
        modelCalls: z.number().int().nonnegative().max(24),
        durationMs: z.number().finite().nonnegative(),
        p50Ms: z.number().finite().nonnegative(),
        p95Ms: z.number().finite().nonnegative(),
      })
      .strict(),
  })
  .strict()
  .superRefine((report, ctx) => {
    const matrix = {
      'contract-review': { surface: 'word', model: true },
      'workbook-close': { surface: 'excel', model: false },
      'board-pack-refresh': { surface: 'powerpoint', model: false },
      'outlook-response': { surface: 'outlook', model: true },
    } as const;
    const counts = { passed: 0, failed: 0, unavailable: 0 };
    for (const attempt of report.attempts) counts[attempt.status]++;
    const expectedStatus =
      counts.failed || !report.sourceFreshnessVerified || !report.approvalRefusalVerified
        ? 'FAIL'
        : counts.unavailable || report.dirty || report.providerScope === 'unavailable'
          ? 'BLOCKED_EXTERNAL'
          : 'PASSED';
    if (report.status !== expectedStatus)
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Workflow status must agree with failed, missing and dirty evidence.',
      });
    const repetitions = new Set(report.attempts.map((attempt) => attempt.repetition));
    if (
      report.attempts.length % 4 !== 0 ||
      Array.from({ length: repetitions.size }, (_, index) => index + 1).some(
        (repetition) => !repetitions.has(repetition),
      ) ||
      [...repetitions].some((repetition) =>
        Object.keys(matrix).some(
          (id) =>
            !report.attempts.some(
              (attempt) => attempt.repetition === repetition && attempt.workflowId === id,
            ),
        ),
      ) ||
      report.attempts.some(
        (attempt) =>
          attempt.surface !== matrix[attempt.workflowId].surface ||
          attempt.modelRequired !== matrix[attempt.workflowId].model,
      )
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message:
          'Every repetition requires the four canonical workflows and their actual model requirements.',
      });
    if (
      (report.providerScope === 'unavailable' && report.identityScope !== 'unavailable') ||
      (report.providerScope === 'replay' && report.identityScope !== 'synthetic-replay') ||
      (report.providerScope === 'live' && report.identityScope !== 'explicit-developer-token')
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Provider and identity observation scopes must agree.',
      });
    if (
      report.attempts.length !== report.summary.requested ||
      counts.passed !== report.summary.passed ||
      counts.failed !== report.summary.failed ||
      counts.unavailable !== report.summary.unavailable ||
      report.attempts.reduce((sum, a) => sum + a.modelCalls, 0) !== report.summary.modelCalls
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Workflow report totals must match every requested attempt.',
      });
    if (
      new Set(report.attempts.map((a) => `${a.workflowId}:${a.repetition}`)).size !==
      report.attempts.length
    )
      ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'Duplicate workflow attempts.' });
    if (
      report.status === 'PASSED' &&
      (report.dirty ||
        report.providerScope === 'unavailable' ||
        report.summary.failed ||
        report.summary.unavailable ||
        !report.sourceFreshnessVerified ||
        !report.approvalRefusalVerified ||
        report.attempts.some(
          (a) =>
            !a.outcomeVerified ||
            (a.modelRequired && (!a.requestRoutingVerified || !a.skillInvocationObserved)),
        ))
    )
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: 'Passing workflow evidence requires observed outcomes, safety and model routing.',
      });
    for (const attempt of report.attempts) {
      if (
        (attempt.status === 'passed' &&
          (!attempt.outcomeVerified ||
            attempt.issueCodes.length ||
            (attempt.modelRequired &&
              (!attempt.modelCalls ||
                !attempt.requestRoutingVerified ||
                !attempt.skillInvocationObserved)))) ||
        (!attempt.modelRequired &&
          (attempt.modelCalls || attempt.requestRoutingVerified || attempt.skillInvocationObserved))
      )
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          message:
            'Attempt outcomes and provider observations must agree with the tested execution path.',
        });
    }
  });
export type TenantWorkflowReport = z.infer<typeof TenantWorkflowReportSchema>;
export type TenantWorkflowAttempt = z.infer<typeof TenantWorkflowAttemptSchema>;
export type TenantWorkflowId = z.infer<typeof TenantWorkflowIdSchema>;
