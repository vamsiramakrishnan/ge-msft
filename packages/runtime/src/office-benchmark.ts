import { z } from 'zod';
import {
  ActuationRequestSchema,
  HostCapabilityProfileSchema,
  OfficeBenchmarkComparisonReportSchema,
  OfficeBenchmarkObservedChecksSchema,
  assessActuationResult,
  contentHash,
  validateActuationResult,
  type ActuationKind,
  type ActuationRequest,
  type OfficeBenchmarkCheck,
  type OfficeBenchmarkComparisonReport,
  type OfficeBenchmarkLeg,
} from '@ge/contracts';
import type { AssistSession, RunCommandsOptions } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import type { OfficeWorkflow } from './office-workflows.js';
import { nativeTraces, type NativeTrace, type NativeTraceStore } from './native-trace.js';

const hash = z.string().regex(/^sha256:[a-f0-9]{64}$/);
const ProbeSnapshotSchema = z
  .object({ scopeHash: hash, sourceHash: hash, artifactHash: hash })
  .strict();
const ProbeOutcomeSchema = ProbeSnapshotSchema.extend({
  checks: OfficeBenchmarkObservedChecksSchema,
}).strict();
const MeterSchema = z
  .object({
    requests: z.number().int().nonnegative(),
    inputBytes: z.number().int().nonnegative(),
    complete: z.boolean(),
  })
  .strict();
export type OfficeBenchmarkInputSnapshot = z.infer<typeof ProbeSnapshotSchema>;
export type OfficeBenchmarkOutcome = z.infer<typeof ProbeOutcomeSchema>;
export interface OfficeBenchmarkProbe {
  /** Stable public code IDs, never document names or content. Hash normalization belongs to this probe. */
  id: string;
  evaluatorId: string;
  input(): Promise<OfficeBenchmarkInputSnapshot>;
  outcome(): Promise<OfficeBenchmarkOutcome>;
}
/** Read-only cumulative counts from the actual provider transport; do not count auth headers. */
export interface OfficeBenchmarkInputMeter {
  read(): { requests: number; inputBytes: number; complete: boolean };
}
export interface OfficeBenchmarkLegInput {
  session: AssistSession;
  bridge: Pick<DocBridge, 'surface' | 'documentKey' | 'getHostProfile'>;
  disposable: true;
  execution: { kind: 'program'; program: string } | { kind: 'workflow'; workflow: OfficeWorkflow };
  expectedEffects: ReadonlyArray<{ kind: ActuationKind; completion: 'verified' | 'review-opened' }>;
  probe: OfficeBenchmarkProbe;
  inputMeter?: OfficeBenchmarkInputMeter;
  options: RunCommandsOptions;
}
export interface OfficeBenchmarkComparisonInput {
  commit: string;
  workflowId: OfficeBenchmarkComparisonReport['workflowId'];
  profileId: string;
  hostScope: OfficeBenchmarkComparisonReport['hostScope'];
  providerScope: OfficeBenchmarkComparisonReport['providerScope'];
  /** Stop other tasks, native capture and provider requests in this process before measuring. */
  confirmExclusiveMeasurement: true;
  baseline: OfficeBenchmarkLegInput;
  candidate: OfficeBenchmarkLegInput;
  signal?: AbortSignal;
}
type Issue = OfficeBenchmarkComparisonReport['issueCodes'][number];
const nativePlatforms = new Set(['PC', 'Mac', 'OfficeOnline', 'iOS', 'Android', 'Universal']);
const workflowSurface = {
  'contract-review': 'word',
  'workbook-close': 'excel',
  'board-pack-refresh': 'powerpoint',
  'outlook-response': 'outlook',
} as const;
let measuring = false;
const issues = (values: Issue[]): Issue[] => [...new Set(values)].slice(0, 32);
async function readProbe<T>(read: () => Promise<T>): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      read(),
      new Promise<never>((_resolve, reject) => {
        timeout = setTimeout(() => reject(new Error('Benchmark read-only probe timed out.')), 2000);
      }),
    ]);
  } finally {
    if (timeout !== undefined) clearTimeout(timeout);
  }
}
const combine = (
  left: OfficeBenchmarkCheck | undefined,
  right: OfficeBenchmarkCheck | undefined,
): OfficeBenchmarkCheck =>
  left === 'failed' || right === 'failed'
    ? 'failed'
    : left === 'passed' && right === 'passed'
      ? 'passed'
      : 'unavailable';
function emptyLeg(input: OfficeBenchmarkLegInput): OfficeBenchmarkLeg {
  return {
    status: 'unavailable',
    durationMs: 0,
    nativeSyncs: 0,
    syncMeasurementComplete: false,
    modelCalls: 0,
    queryBytes: 0,
    queryBytesComplete: false,
    inferenceInputBytes: null,
    inferenceInputComplete: false,
    effects: [],
    probeId: input.probe.id,
    evaluatorId: input.probe.evaluatorId,
    issueCodes: [],
  };
}
function profile(
  leg: OfficeBenchmarkLegInput,
): ReturnType<typeof HostCapabilityProfileSchema.parse> | undefined {
  const raw = leg.bridge.getHostProfile?.();
  if (!raw) return undefined;
  const parsed = HostCapabilityProfileSchema.parse(raw);
  if (Object.keys(parsed.requirements).length > 128)
    throw new Error('Host profile exceeds the benchmark metadata limit.');
  return parsed;
}
function validateInput(input: OfficeBenchmarkComparisonInput): void {
  if (
    !/^[a-f0-9]{40}$/.test(input.commit) ||
    !/^[a-z][a-z0-9._-]{0,63}$/.test(input.profileId) ||
    input.confirmExclusiveMeasurement !== true
  )
    throw new Error(
      'Use an exact commit, public profile ID and explicit exclusive measurement confirmation.',
    );
  if (input.baseline.session === input.candidate.session)
    throw new Error('Use separate disposable sessions for baseline and candidate.');
  for (const leg of [input.baseline, input.candidate]) {
    if (
      leg.disposable !== true ||
      leg.expectedEffects.length < 1 ||
      leg.expectedEffects.length > 32 ||
      !/^[a-z][a-z0-9._-]{0,63}$/.test(leg.probe.id) ||
      !/^[a-z][a-z0-9._-]{0,63}$/.test(leg.probe.evaluatorId)
    )
      throw new Error(
        'Each leg needs a disposable target, bounded expected effects and public probe/evaluator IDs.',
      );
    if (
      leg.execution.kind === 'program' &&
      new TextEncoder().encode(leg.execution.program).byteLength > 256 * 1024
    )
      throw new Error('The benchmark program exceeds 256 KiB.');
  }
}
/**
 * Measures existing execution, never dispatches a host operation itself. Input/output probes run
 * outside the execution timing/sync window. They must be trusted bounded native reads with an
 * explicit evaluator; their checks are observations, not default safety-success counters.
 */
export async function runOfficeBenchmarkComparison(
  input: OfficeBenchmarkComparisonInput,
  traces: NativeTraceStore = nativeTraces,
): Promise<OfficeBenchmarkComparisonReport> {
  validateInput(input);
  if (measuring) throw new Error('An Office comparison is already measuring this process.');
  measuring = true;
  try {
    const baseline = await runLeg(input, input.baseline, traces);
    const candidate = await runLeg(input, input.candidate, traces);
    const checks: OfficeBenchmarkComparisonReport['checks'] = {
      sameInput:
        baseline.sourceHash &&
        candidate.sourceHash &&
        baseline.initialArtifactHash &&
        candidate.initialArtifactHash
          ? baseline.sourceHash === candidate.sourceHash &&
            baseline.initialArtifactHash === candidate.initialArtifactHash &&
            baseline.probeId === candidate.probeId &&
            baseline.evaluatorId === candidate.evaluatorId
            ? 'passed'
            : 'failed'
          : 'unavailable',
      artifactsMatch:
        baseline.artifactHash && candidate.artifactHash
          ? baseline.artifactHash === candidate.artifactHash
            ? 'passed'
            : 'failed'
          : 'unavailable',
      sourceFreshness:
        baseline.issueCodes.includes('source-changed') ||
        candidate.issueCodes.includes('source-changed')
          ? 'failed'
          : baseline.artifactHash && candidate.artifactHash
            ? 'passed'
            : 'unavailable',
      targetIntegrity: combine(baseline.checks?.targetIntegrity, candidate.checks?.targetIntegrity),
      approvalEnforced: combine(
        baseline.checks?.approvalEnforced,
        candidate.checks?.approvalEnforced,
      ),
      completionTruth: combine(baseline.checks?.completionTruth, candidate.checks?.completionTruth),
      freshnessRefusal: combine(
        baseline.checks?.freshnessRefusal,
        candidate.checks?.freshnessRefusal,
      ),
    };
    const failureCodes: Issue[] = [...baseline.issueCodes, ...candidate.issueCodes];
    if (checks.sameInput === 'failed') failureCodes.push('different-input');
    if (checks.artifactsMatch === 'failed') failureCodes.push('different-artifact');
    const passed =
      baseline.status === 'passed' &&
      candidate.status === 'passed' &&
      Object.values(checks).every((value) => value === 'passed');
    const unavailable =
      baseline.status === 'unavailable' ||
      candidate.status === 'unavailable' ||
      Object.values(checks).some((value) => value === 'unavailable');
    const failed =
      baseline.status === 'failed' ||
      candidate.status === 'failed' ||
      Object.values(checks).some((value) => value === 'failed');
    return OfficeBenchmarkComparisonReportSchema.parse({
      version: 1,
      id: crypto.randomUUID(),
      commit: input.commit,
      workflowId: input.workflowId,
      profileId: input.profileId,
      hostScope: input.hostScope,
      providerScope: input.providerScope,
      status: passed ? 'passed' : failed ? 'failed' : unavailable ? 'unavailable' : 'failed',
      baseline,
      candidate,
      checks,
      measurements: {
        syncReduction:
          passed &&
          baseline.syncMeasurementComplete &&
          candidate.syncMeasurementComplete &&
          baseline.nativeSyncs > 0
            ? 1 - candidate.nativeSyncs / baseline.nativeSyncs
            : null,
        inferenceInputReduction:
          passed &&
          baseline.inferenceInputComplete &&
          candidate.inferenceInputComplete &&
          baseline.inferenceInputBytes !== null &&
          candidate.inferenceInputBytes !== null &&
          baseline.inferenceInputBytes > 0
            ? 1 - candidate.inferenceInputBytes / baseline.inferenceInputBytes
            : null,
      },
      issueCodes: issues(failureCodes),
    });
  } finally {
    measuring = false;
  }
}

async function runLeg(
  comparison: OfficeBenchmarkComparisonInput,
  input: OfficeBenchmarkLegInput,
  traces: NativeTraceStore,
): Promise<OfficeBenchmarkLeg> {
  const result = emptyLeg(input);
  let before: z.infer<typeof ProbeSnapshotSchema>;
  let startHost: ReturnType<typeof profile>;
  const startKey = input.bridge.documentKey?.();
  try {
    startHost = profile(input);
    if (startHost) result.host = startHost;
    if (input.bridge.surface !== workflowSurface[comparison.workflowId]) {
      result.issueCodes.push('wrong-surface');
      return result;
    }
    if (
      comparison.hostScope === 'native' &&
      (!startHost ||
        !nativePlatforms.has(startHost.platform) ||
        !startHost.build ||
        startHost.build === 'unknown' ||
        startHost.mode === 'unknown')
    ) {
      result.issueCodes.push('unavailable-host');
      return result;
    }
    if (!startKey) {
      result.issueCodes.push('unavailable-host');
      return result;
    }
    if (!input.options.approvePlan && !input.options.approveWrite) {
      result.issueCodes.push('missing-approval');
      return result;
    }
    comparison.signal?.throwIfAborted();
    const binding = await input.session.applicationContext();
    if (
      binding.surface !== input.bridge.surface ||
      binding.documentId !==
        (await contentHash({ surface: input.bridge.surface, documentKey: startKey }))
    ) {
      result.issueCodes.push('probe-changed-scope');
      return result;
    }
    before = ProbeSnapshotSchema.parse(await readProbe(() => input.probe.input()));
    if (input.bridge.documentKey?.() !== startKey) {
      result.issueCodes.push('probe-changed-scope');
      return result;
    }
    result.scopeHash = before.scopeHash;
    result.sourceHash = before.sourceHash;
    result.initialArtifactHash = before.artifactHash;
  } catch {
    result.issueCodes.push(comparison.signal?.aborted ? 'cancelled' : 'probe-failed');
    return result;
  }
  const priorRuns = new Set(input.session.executions.list().map((run) => run.taskId));
  const requests = new Map<string, ActuationRequest>();
  const observed: NativeTrace[] = [];
  let overflow = false;
  const offRequest = input.session.hooks.register({
    id: `benchmark-${crypto.randomUUID()}`,
    on: 'effect:before',
    mode: 'observe',
    handle: (payload) => {
      const request = ActuationRequestSchema.parse(payload.request);
      requests.set(request.changeId, request);
    },
  });
  const offTrace = traces.subscribe((trace) => {
    if (observed.length < 10000) observed.push(trace);
    else overflow = true;
  });
  const receiptIssues: Issue[] = [];
  const receipts: OfficeBenchmarkLeg['effects'] = [];
  let finished = false;
  let executionError = false;
  let startMeter: z.infer<typeof MeterSchema> | undefined;
  try {
    if (input.inputMeter) startMeter = MeterSchema.parse(input.inputMeter.read());
  } catch {
    receiptIssues.push('missing-metrics');
  }
  const started = performance.now();
  try {
    let approvedEffects = 0;
    const options = {
      ...input.options,
      ...(comparison.signal ? { signal: comparison.signal } : {}),
      ...(input.options.approvePlan
        ? {
            approvePlan: async (
              effects: Parameters<NonNullable<RunCommandsOptions['approvePlan']>>[0],
            ) => {
              const expected = input.expectedEffects.slice(
                approvedEffects,
                approvedEffects + effects.length,
              );
              if (
                expected.length !== effects.length ||
                effects.some((effect, index) => effect.request.kind !== expected[index]?.kind)
              )
                return false;
              const approved = await input.options.approvePlan!(effects);
              if (approved) approvedEffects += effects.length;
              return approved;
            },
          }
        : {
            approveWrite: async (request: ActuationRequest) => {
              if (request.kind !== input.expectedEffects[approvedEffects]?.kind) return false;
              const approved = await input.options.approveWrite!(request);
              if (approved) approvedEffects++;
              return approved;
            },
          }),
    };
    const events =
      input.execution.kind === 'program'
        ? input.session.runCommandProgram(input.execution.program, options)
        : input.session.runOfficeWorkflow(input.execution.workflow, options);
    for await (const event of events) {
      if (event.type === 'done' && 'turn' in event) finished = !event.incomplete;
      if (event.type === 'error' || event.type === 'capped' || event.type === 'needs-user')
        executionError = true;
      if (event.type !== 'write-result') continue;
      const request = requests.get(event.changeId);
      if (!request || event.result.changeId !== event.changeId) {
        receiptIssues.push('receipt-mismatch');
        continue;
      }
      const receipt = validateActuationResult(request, event.result);
      const assessment = assessActuationResult(receipt);
      receipts.push({
        kind: receipt.kind,
        assessment:
          receipt.ok && receipt.completion === 'review-opened' && !receipt.recoveryPending
            ? 'review-opened'
            : assessment,
      });
      if (receipts.length > 64) {
        receiptIssues.push('receipt-mismatch');
        break;
      }
    }
  } catch {
    executionError = true;
    receiptIssues.push(comparison.signal?.aborted ? 'cancelled' : 'execution-failed');
  } finally {
    result.durationMs = performance.now() - started;
    offRequest();
    offTrace();
  }
  result.effects = receipts.slice(0, 64);
  if (executionError) receiptIssues.push('execution-failed');
  if (!finished) receiptIssues.push('incomplete');
  const runs = input.session.executions.list().filter((run) => !priorRuns.has(run.taskId));
  result.modelCalls = runs.reduce((sum, run) => sum + run.modelTurns, 0);
  result.queryBytes = runs.reduce((sum, run) => sum + (run.metrics?.queryBytes ?? 0), 0);
  result.queryBytesComplete =
    runs.length === 1 && runs.every((run) => run.modelTurns === 0 || run.metrics !== undefined);
  if (runs.length !== 1) receiptIssues.push('task-overlap');
  if (runs.some((run) => run.status !== 'completed'))
    receiptIssues.push(comparison.signal?.aborted ? 'cancelled' : 'execution-failed');
  if (comparison.providerScope === 'none' && result.modelCalls !== 0)
    receiptIssues.push('invalid-origin');
  if (result.modelCalls === 0) {
    result.inferenceInputBytes = 0;
    result.inferenceInputComplete = true;
  } else if (startMeter && input.inputMeter) {
    try {
      const end = MeterSchema.parse(input.inputMeter.read());
      const requests = end.requests - startMeter.requests;
      const bytes = end.inputBytes - startMeter.inputBytes;
      if (end.complete && startMeter.complete && requests === result.modelCalls && bytes >= 0) {
        result.inferenceInputBytes = bytes;
        result.inferenceInputComplete = true;
      } else receiptIssues.push('missing-metrics');
    } catch {
      receiptIssues.push('missing-metrics');
    }
  }
  const foreign = observed.some(
    (trace) =>
      trace.surface !== input.bridge.surface ||
      (trace.stage === 'dispatch' && (!trace.changeId || !requests.has(trace.changeId))),
  );
  result.nativeSyncs = observed.reduce(
    (sum, trace) => sum + (trace.stage === 'batch' ? trace.syncs : 0),
    0,
  );
  result.syncMeasurementComplete = !foreign && !overflow;
  if (foreign) receiptIssues.push('trace-contamination');
  if (overflow) receiptIssues.push('trace-limit');
  if (
    receipts.length !== input.expectedEffects.length ||
    receipts.some((receipt, index) => receipt.kind !== input.expectedEffects[index]?.kind)
  )
    receiptIssues.push('receipt-mismatch');
  else if (
    receipts.some(
      (receipt, index) => receipt.assessment !== input.expectedEffects[index]?.completion,
    )
  )
    receiptIssues.push('unverified-effect');
  try {
    comparison.signal?.throwIfAborted();
    const after = ProbeOutcomeSchema.parse(await readProbe(() => input.probe.outcome()));
    if (after.scopeHash !== before.scopeHash || input.bridge.documentKey?.() !== startKey)
      receiptIssues.push('probe-changed-scope');
    if (after.sourceHash !== before.sourceHash) receiptIssues.push('source-changed');
    if (JSON.stringify(profile(input)) !== JSON.stringify(startHost))
      receiptIssues.push('host-changed');
    result.artifactHash = after.artifactHash;
    result.checks = after.checks;
    const checkCodes = {
      targetIntegrity: 'target-check',
      approvalEnforced: 'approval-check',
      completionTruth: 'completion-check',
      freshnessRefusal: 'freshness-check',
    } as const;
    for (const [key, code] of Object.entries(checkCodes))
      if (after.checks[key as keyof typeof after.checks] !== 'passed') receiptIssues.push(code);
  } catch {
    receiptIssues.push(comparison.signal?.aborted ? 'cancelled' : 'probe-failed');
  }
  result.issueCodes = issues([...result.issueCodes, ...receiptIssues]);
  result.status = result.issueCodes.length === 0 ? 'passed' : 'failed';
  return result;
}
