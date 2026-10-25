import { z } from 'zod';
import type { ActuationResult, Surface, CapabilityManifest, SseEvent } from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import type { AssistSession, RunCommandsOptions, CommandLoopEvent } from './assist-session.js';
import { nativeTraces } from './native-trace.js';

export const NATIVE_FIXTURE_VERSION = 1;
export interface NativeFixture {
  id: string;
  surface: Surface;
  program: string;
  expectedKind: ActuationResult['kind'];
  expectedText?: string;
  cleanup: string;
}
/** Run only in a disposable file/draft. The normal approval card shows every fixture mutation. */
export function nativeFixture(surface: Surface, targetRange?: string): NativeFixture {
  const nonce = `GE_CERT_${crypto.randomUUID().replace(/-/g, '')}`;
  if (surface === 'word')
    return {
      id: 'word-tracked-edit',
      surface,
      program: `/insert-text position=end text="${nonce}_OLD"\nsuggest "${nonce}_OLD" => "${nonce}_NEW"`,
      expectedKind: 'tracked-change',
      expectedText: `${nonce}_NEW`,
      cleanup: 'Reject or remove the fixture tracked edits in the disposable document.',
    };
  if (surface === 'excel') {
    if (!targetRange || /[\r\n`]/.test(targetRange))
      throw new Error('Supply an explicit disposable workbook target range.');
    return {
      id: 'excel-grid-readback',
      surface,
      program: `grid ${targetRange} = "Metric\\tValue\\n${nonce}\\t12.34"\nfinish when=verified`,
      expectedKind: 'write-cells',
      cleanup: 'Use the reviewed Recovery & undo action to restore the fixture range.',
    };
  }
  if (surface === 'powerpoint')
    return {
      id: 'powerpoint-slide-readback',
      surface,
      program: `slide "${nonce}" "Native text readback"`,
      expectedKind: 'insert-slide',
      expectedText: nonce,
      cleanup: 'Remove the fixture slide from the disposable presentation.',
    };
  if (surface === 'outlook')
    return {
      id: 'outlook-review-form',
      surface,
      program: `compose "${nonce}" "Native certification: review form only."`,
      expectedKind: 'create-mail',
      cleanup: 'Discard the fixture draft without sending it.',
    };
  throw new Error('No native certification fixture is promoted for this surface.');
}
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
  bridge: Pick<DocBridge, 'surface' | 'readRange' | 'getHostProfile'>,
  fixture: NativeFixture,
  commit: string,
  options: RunCommandsOptions,
  observe?: (event: SseEvent | CommandLoopEvent) => void,
): Promise<NativeCertificationReport> {
  if (!/^[a-f0-9]{40}$/.test(commit) || bridge.surface !== fixture.surface)
    throw new Error('Certification requires the exact tested commit and matching native surface.');
  const started = performance.now();
  const profile = bridge.getHostProfile?.();
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
  if (!profile || profile.platform === 'unknown' || profile.build === 'unknown') return report;
  const prior = new Set(nativeTraces.list().map((trace) => trace.id));
  const receipts: ActuationResult[] = [];
  let executionError = false;
  let finished = false;
  try {
    for await (const event of session.runCommandProgram(fixture.program, options)) {
      try {
        observe?.(event);
      } catch {
        /* observational UI only */
      }
      if (event.type === 'done' && 'turn' in event) finished = !event.incomplete;
      if (event.type === 'write-result') receipts.push(event.result);
      if (event.type === 'error' || event.type === 'capped' || event.type === 'needs-user')
        executionError = true;
    }
    const effect = receipts.find((receipt) => receipt.kind === fixture.expectedKind);
    report.checks.finished = finished;
    report.checks.execution =
      !executionError &&
      receipts.length > 0 &&
      receipts.every((receipt) => receipt.ok && !receipt.recoveryPending && !receipt.error);
    report.checks.expectedEffect = Boolean(effect?.ok);
    if (fixture.surface === 'word' || fixture.surface === 'excel')
      report.checks.postcondition = effect?.verification?.status === 'verified';
    if (fixture.surface === 'outlook')
      report.checks.postcondition = effect?.completion === 'review-opened';
    if (fixture.surface === 'powerpoint') {
      const contexts =
        effect?.location && bridge.readRange ? await bridge.readRange(effect.location) : [];
      report.checks.postcondition = contexts.some(
        (context) =>
          context.value.as === 'text' && context.value.text.includes(fixture.expectedText!),
      );
    }
    report.status = Object.values(report.checks).every(Boolean) ? 'passed' : 'failed';
  } catch {
    report.status = 'failed';
    report.checks.finished = finished;
    report.checks.execution = false;
  }
  report.durationMs = performance.now() - started;
  report.nativeSyncs = nativeTraces
    .list()
    .filter(
      (trace) =>
        !prior.has(trace.id) && trace.surface === fixture.surface && trace.stage === 'batch',
    )
    .reduce((sum, trace) => sum + trace.syncs, 0);
  report.effects = receipts.map((receipt) => ({
    kind: receipt.kind,
    ok: receipt.ok,
    ...(receipt.verification ? { verification: receipt.verification.status } : {}),
    ...(receipt.completion ? { completion: receipt.completion } : {}),
  }));
  return NativeCertificationReportSchema.parse(report);
}

/** The same capability intersection used by clients; certification never promotes a missing handler. */
export function nativeCertificationCoverage(manifest: CapabilityManifest): string[] {
  return manifest.actuations
    .filter((entry) =>
      ['tracked-change', 'write-cells', 'insert-slide', 'create-mail'].includes(entry.kind),
    )
    .map((entry) => entry.kind);
}
