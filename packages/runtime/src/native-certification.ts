import { z } from 'zod';
import {
  ActuationRequestSchema,
  capabilityModule,
  validateActuationResult,
  type ActuationRequest,
  type ActuationResult,
  type CapabilityManifest,
  type HostCapabilityProfile,
  type SseEvent,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import type { AssistSession, RunCommandsOptions, CommandLoopEvent } from './assist-session.js';
import { nativeTraces } from './native-trace.js';
import { capabilitySupport } from './host-profile.js';
import { isNativeFixture, NATIVE_FIXTURE_VERSION, type NativeFixture } from './native-fixtures.js';
export * from './native-fixtures.js';
export const NativeCertificationReportSchema = z
  .object({
    version: z.literal(1),
    fixtureVersion: z.literal(NATIVE_FIXTURE_VERSION),
    fixtureId: z.string(),
    commit: z.string().regex(/^[a-f0-9]{40}$/),
    surface: z.enum(['word', 'excel', 'powerpoint', 'outlook', 'onenote', 'teams']),
    host: z
      .object({
        platform: z.string(),
        build: z.string(),
        mode: z.string(),
        requirements: z.record(z.boolean()),
      })
      .optional(),
    startedAt: z.string().datetime(),
    durationMs: z.number().nonnegative(),
    status: z.enum(['passed', 'failed', 'unavailable']),
    checks: z.record(z.boolean()),
    nativeSyncs: z.number().int().nonnegative(),
    effects: z.array(
      z.object({
        kind: z.string(),
        ok: z.boolean(),
        verification: z.string().optional(),
        completion: z.string().optional(),
      }),
    ),
  })
  .strict();
export type NativeCertificationReport = z.infer<typeof NativeCertificationReportSchema>;
/** Uses the existing executor, including approval, policy, freshness checks, recovery and journal. */
export async function certifyNativeFixture(
  session: Pick<AssistSession, 'runCommandProgram'>,
  bridge: Pick<DocBridge, 'surface' | 'readRange' | 'getHostProfile' | 'documentKey'> &
    Partial<Pick<DocBridge, 'getCapabilities'>>,
  fixture: NativeFixture,
  commit: string,
  options: RunCommandsOptions,
  observe?: (event: SseEvent | CommandLoopEvent) => void,
): Promise<NativeCertificationReport> {
  if (!isNativeFixture(fixture))
    throw new Error('Use an unchanged fixture created by the native fixture factory.');
  if (!/^[a-f0-9]{40}$/.test(commit) || bridge.surface !== fixture.surface)
    throw new Error('Certification requires the exact tested commit and matching native surface.');
  const started = performance.now();
  let profile: HostCapabilityProfile | undefined;
  try {
    const observed = bridge.getHostProfile?.();
    profile = observed ? structuredClone(observed) : undefined;
  } catch {
    /* failed native discovery is unavailable */
  }
  const report: NativeCertificationReport = {
    version: 1,
    fixtureVersion: NATIVE_FIXTURE_VERSION,
    fixtureId: fixture.id,
    commit,
    surface: fixture.surface,
    startedAt: new Date().toISOString(),
    durationMs: 0,
    status: 'unavailable',
    checks: {},
    nativeSyncs: 0,
    effects: [],
    ...(profile
      ? {
          host: {
            platform: profile.platform,
            build: profile.build,
            mode: profile.mode,
            requirements: profile.requirements,
          },
        }
      : {}),
  };
  let documentKey: string | undefined;
  try {
    documentKey = bridge.documentKey?.();
  } catch {
    /* no certified identity */
  }
  report.checks.nativeHost = Boolean(
    profile &&
    profile.surface === fixture.surface &&
    profile.platform &&
    profile.build &&
    !/unknown|simulator|test-only|fake/i.test(profile.platform + ' ' + profile.build),
  );
  report.checks.documentIdentity = Boolean(documentKey);
  report.checks.mode = Boolean(profile && fixture.modes.includes(profile.mode));
  report.checks.support = Boolean(
    profile &&
    fixture.expectedEffects.every(
      (kind) => capabilitySupport(fixture.surface, kind, profile).supported,
    ),
  );
  if (!profile || !Object.values(report.checks).every(Boolean)) return report;
  let manifest: CapabilityManifest | undefined;
  try {
    manifest = bridge.getCapabilities
      ? await boundedRead(() => Promise.resolve(bridge.getCapabilities!()))
      : undefined;
  } catch {
    report.checks.support = false;
    return report;
  }
  if (
    manifest &&
    (manifest.surface !== fixture.surface ||
      fixture.expectedEffects.some(
        (kind) => !manifest.actuations.some((entry) => entry.kind === kind),
      ))
  )
    return report;
  let nativeSyncs = 0;
  const stopObserving = nativeTraces.subscribe((trace) => {
    if (trace.surface === fixture.surface && trace.stage === 'batch') nativeSyncs += trace.syncs;
  });
  const receipts: ActuationResult[] = [];
  let executionError = false;
  let finished = false;
  let approvalsValid = true;
  let correlated = true;
  const approved = new Map<string, ActuationRequest>();
  const profileBefore = JSON.stringify(profile);
  const phaseEffects = fixture.setupProgram
    ? [fixture.expectedEffects.slice(0, 1), fixture.expectedEffects.slice(1)]
    : [fixture.expectedEffects];
  const programs = fixture.setupProgram
    ? [fixture.setupProgram, fixture.program]
    : [fixture.program];
  try {
    for (const [phase, program] of programs.entries()) {
      if (phase > 0 && (!finished || executionError || !correlated || !approvalsValid)) break;
      finished = false;
      const expected = phaseEffects[phase]!;
      const offset = receipts.length;
      const guarded: RunCommandsOptions = {
        ...options,
        beforeDispatch: async () => {
          if (
            bridge.documentKey?.() !== documentKey ||
            JSON.stringify(bridge.getHostProfile?.()) !== profileBefore
          )
            throw new Error('Native fixture host or document changed.');
          await options.beforeDispatch?.();
          if (
            bridge.documentKey?.() !== documentKey ||
            JSON.stringify(bridge.getHostProfile?.()) !== profileBefore
          )
            throw new Error('Native fixture host or document changed during source validation.');
        },
        approvePlan: async (effects) => {
          const requests = effects.map((effect) => ActuationRequestSchema.parse(effect.request));
          const valid =
            requests.length === expected.length &&
            new Set(requests.map((request) => request.changeId)).size === requests.length &&
            requests.every(
              (request, index) =>
                request.kind === expected[index] &&
                request.surface === fixture.surface &&
                !approved.has(request.changeId) &&
                capabilityModule(fixture.surface, request.kind)?.params.safeParse(request.params)
                  .success === true,
            );
          if (!valid) {
            approvalsValid = false;
            return false;
          }
          let accepted = false;
          if (options.approvePlan) accepted = await options.approvePlan(effects);
          else if (options.approveWrite) {
            accepted = true;
            for (const request of requests) {
              if (!(await options.approveWrite(structuredClone(request)))) {
                accepted = false;
                break;
              }
            }
          }
          if (accepted) for (const request of requests) approved.set(request.changeId, request);
          return accepted;
        },
      };
      for await (const event of session.runCommandProgram(program, guarded)) {
        try {
          observe?.(event);
        } catch {
          /* observational UI only */
        }
        if (event.type === 'done' && 'turn' in event) finished = !event.incomplete;
        if (event.type === 'write-result') {
          const request = approved.get(event.changeId);
          if (
            !request ||
            event.changeId !== event.result.changeId ||
            request.kind !== event.result.kind
          )
            correlated = false;
          const receipt = request ? validateActuationResult(request, event.result) : event.result;
          receipts.push(receipt);
        }
        if (event.type === 'error' || event.type === 'capped' || event.type === 'needs-user')
          executionError = true;
      }
      const phaseReceipts = receipts.slice(offset);
      if (
        phaseReceipts.length !== expected.length ||
        phaseReceipts.some(
          (receipt, index) =>
            receipt.kind !== expected[index] ||
            !receipt.ok ||
            receipt.recoveryPending ||
            receipt.trackingModeRestoreFailed ||
            receipt.degraded ||
            receipt.error,
        )
      )
        executionError = true;
      if (
        bridge.documentKey?.() !== documentKey ||
        JSON.stringify(bridge.getHostProfile?.()) !== profileBefore
      )
        executionError = true;
    }
    const effect = receipts.find((receipt) => receipt.kind === fixture.expectedKind);
    report.checks.finished = finished;
    report.checks.approved = approvalsValid && approved.size === fixture.expectedEffects.length;
    report.checks.correlated = correlated;
    report.checks.expectedEffects =
      receipts.length === fixture.expectedEffects.length &&
      receipts.every((receipt, index) => receipt.kind === fixture.expectedEffects[index]);
    report.checks.scopeStable = bridge.documentKey?.() === documentKey;
    report.checks.hostStable = JSON.stringify(bridge.getHostProfile?.()) === profileBefore;
    report.checks.execution =
      !executionError &&
      receipts.length > 0 &&
      receipts.every(
        (receipt) =>
          receipt.ok &&
          !receipt.recoveryPending &&
          !receipt.trackingModeRestoreFailed &&
          !receipt.degraded &&
          !receipt.error,
      );
    report.checks.expectedEffect = Boolean(effect?.ok);
    if (fixture.verifiedEffects)
      report.checks.postcondition = fixture.verifiedEffects.every(
        (index) => receipts[index]?.verification?.status === 'verified',
      );
    if (fixture.reviewOpened) report.checks.postcondition = effect?.completion === 'review-opened';
    if (fixture.surface === 'powerpoint' && !fixture.verifiedEffects) {
      const contexts =
        effect?.location && bridge.readRange
          ? await boundedRead(() => bridge.readRange!(effect.location!))
          : [];
      report.checks.postcondition =
        contexts.length <= 16 &&
        new TextEncoder().encode(JSON.stringify(contexts)).byteLength <= 64 * 1024 &&
        contexts.some(
          (context) =>
            context.value.as === 'text' && context.value.text.includes(fixture.expectedText!),
        );
    }
    report.checks.scopeStable &&= bridge.documentKey?.() === documentKey;
    report.checks.hostStable &&= JSON.stringify(bridge.getHostProfile?.()) === profileBefore;
    report.status = Object.values(report.checks).every(Boolean) ? 'passed' : 'failed';
  } catch {
    report.status = 'failed';
    report.checks.finished = finished;
    report.checks.execution = false;
  } finally {
    stopObserving();
  }
  report.durationMs = performance.now() - started;
  report.nativeSyncs = nativeSyncs;
  report.effects = receipts.map((receipt) => ({
    kind: receipt.kind,
    ok: receipt.ok,
    ...(receipt.verification ? { verification: receipt.verification.status } : {}),
    ...(receipt.completion ? { completion: receipt.completion } : {}),
  }));
  return NativeCertificationReportSchema.parse(report);
}

async function boundedRead<T>(read: () => Promise<T>): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      read(),
      new Promise<T>((_resolve, reject) => {
        timer = setTimeout(() => reject(new Error('Native readback timed out.')), 5000);
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

/** The same capability intersection used by clients; certification never promotes a missing handler. */
export function nativeCertificationCoverage(manifest: CapabilityManifest): string[] {
  return manifest.actuations
    .filter((entry) =>
      [
        'tracked-change',
        'write-cells',
        'insert-slide',
        'create-mail',
        'insert-table',
        'create-table',
        'insert-pivot',
        'filter-range',
        'set-data-validation',
        'apply-slide-layout',
        'set-subject',
        'set-body',
      ].includes(entry.kind),
    )
    .map((entry) => entry.kind);
}
