// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { TenantWorkflowReportSchema } from '@ge/contracts';
import {
  configuredTenantWorkflowProvider,
  type TenantWorkflowProvider,
} from './tenant-workflow-env.js';
import { runFlagshipTenantSuite } from './tenant-workflows.js';

const commit = 'a'.repeat(40);
function replay(
  options: { invocation?: boolean; routeReply?: 'bad'; quota?: boolean; wrongAgent?: boolean } = {},
) {
  const fetchImpl = vi.fn(async (_input, init) => {
    if (options.quota) return new Response('{}', { status: 429 });
    const raw = String(init?.body);
    const text =
      options.routeReply === 'bad'
        ? '```cmd\ndone "Cannot complete."\n```'
        : raw.includes('specified contract scope')
          ? '```cmd\nsuggest "Payment is due in 45 days." => "Payment is due in 30 days."\nfinish when=verified\n```'
          : '```cmd\nmail "<p>The security review is open. Mira will update by Friday.</p>"\ndone "A review form is open."\n```';
    return new Response(
      JSON.stringify([
        {
          ...(options.invocation === false
            ? {}
            : {
                invokedSkills: [
                  {
                    name: options.wrongAgent ? 'another-agent' : 'synthetic-commander',
                    displayName: 'synthetic-commander',
                  },
                ],
              }),
          answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
        },
      ]),
      { headers: { 'Content-Type': 'application/json' } },
    );
  });
  const provider: TenantWorkflowProvider = {
    assistant: { project: 'synthetic-test', location: 'eu', engine: 'synthetic-test' },
    commander: { label: 'synthetic-commander', resource: 'synthetic-commander' },
    getAccessToken: async () => 'synthetic-replay-token',
    fetchImpl,
  };
  return { provider, fetchImpl };
}
describe('metadata-only flagship workflow suite', () => {
  it('runs all four real runtime/bridge paths and records labelled replay evidence', async () => {
    const p = replay();
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: false,
      provider: p.provider,
      providerScope: 'replay',
    });
    expect(report.status, JSON.stringify(report)).toBe('PASSED');
    expect(report.hostScope).toBe('simulated');
    expect(report.providerScope).toBe('replay');
    expect(report.sourceFreshnessVerified).toBe(true);
    expect(report.approvalRefusalVerified).toBe(true);
    expect(report.summary.passed).toBe(4);
    expect(report.summary.modelCalls).toBeGreaterThanOrEqual(2);
    expect(
      report.attempts
        .filter((attempt) => !attempt.modelRequired)
        .every((attempt) => attempt.modelCalls === 0 && !attempt.skillInvocationObserved),
    ).toBe(true);
    const serialized = JSON.stringify(report);
    expect(serialized).not.toContain('synthetic-replay-token');
    expect(serialized).not.toContain('Payment is due');
    expect(serialized).not.toContain('9007199254740993');
  }, 60000);
  it('runs deterministic cases but blocks unavailable model evidence without a tenant', async () => {
    const report = await runFlagshipTenantSuite({ commit, dirty: false });
    expect(report.status, JSON.stringify(report)).toBe('BLOCKED_EXTERNAL');
    expect(report.summary).toMatchObject({ passed: 2, unavailable: 2, modelCalls: 0 });
    expect(
      report.attempts
        .filter((attempt) => attempt.modelRequired)
        .every((attempt) => attempt.issueCodes.includes('configuration_missing')),
    ).toBe(true);
  }, 60000);
  it('does not promote requested skill mounting to observed invocation', async () => {
    const p = replay({ invocation: false });
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: false,
      provider: p.provider,
      providerScope: 'replay',
    });
    expect(report.status).toBe('BLOCKED_EXTERNAL');
    expect(
      report.attempts
        .filter((attempt) => attempt.modelRequired)
        .every(
          (attempt) =>
            attempt.requestRoutingVerified &&
            !attempt.skillInvocationObserved &&
            attempt.status === 'unavailable',
        ),
    ).toBe(true);
  }, 60000);
  it('fails wrong outcomes and bounds repeated model calls without persisting raw provider errors', async () => {
    const p = replay({ routeReply: 'bad' });
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: false,
      provider: p.provider,
      providerScope: 'replay',
      repetitions: 2,
      maxModelCalls: 1,
    });
    expect(report.status).toBe('FAIL');
    expect(report.summary.requested).toBe(8);
    expect(p.fetchImpl).toHaveBeenCalledOnce();
    expect(report.attempts.some((attempt) => attempt.issueCodes.includes('query_budget'))).toBe(
      true,
    );
  }, 60000);
  it('requires a clean exact commit before spending live calls and rejects incomplete report matrices', async () => {
    const p = replay();
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: true,
      provider: p.provider,
      providerScope: 'live',
    });
    expect(report.status).toBe('BLOCKED_EXTERNAL');
    expect(p.fetchImpl).not.toHaveBeenCalled();
    expect(
      TenantWorkflowReportSchema.safeParse({ ...report, attempts: report.attempts.slice(0, 3) })
        .success,
    ).toBe(false);
    expect(TenantWorkflowReportSchema.safeParse({ ...report, hostScope: 'native' }).success).toBe(
      false,
    );
  }, 60000);
  it('stops spending the provider pool after quota refusal and leaves native acceptance unavailable', async () => {
    const p = replay({ quota: true });
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: false,
      provider: p.provider,
      providerScope: 'replay',
    });
    expect(report.status).toBe('BLOCKED_EXTERNAL');
    expect(report.summary).toMatchObject({ failed: 0, passed: 2, unavailable: 2, modelCalls: 1 });
    expect(p.fetchImpl).toHaveBeenCalledOnce();
  }, 60000);
});

describe('invocation identity observation', () => {
  it('requires the configured invoked agent even if another agent reuses its display label', async () => {
    const p = replay({ wrongAgent: true });
    const report = await runFlagshipTenantSuite({
      commit,
      dirty: false,
      provider: p.provider,
      providerScope: 'replay',
    });
    expect(report.status).toBe('BLOCKED_EXTERNAL');
    expect(
      report.attempts
        .filter((attempt) => attempt.modelRequired)
        .every((attempt) => !attempt.skillInvocationObserved && attempt.outcomeVerified),
    ).toBe(true);
  }, 60000);
});

describe('explicit tenant configuration', () => {
  it('does not invent credentials, region or an implicit login', async () => {
    expect(configuredTenantWorkflowProvider({})).toBeUndefined();
    expect(
      configuredTenantWorkflowProvider({
        GE_PROJECT: 'synthetic-test',
        GE_ENGINE: 'engine-test',
        GE_SURFACE_COMMANDER_SKILL: 'commander',
        GE_ACCESS_TOKEN: 'synthetic',
      }),
    ).toBeUndefined();
    const provider = configuredTenantWorkflowProvider({
      GE_PROJECT: 'synthetic-test',
      GE_LOCATION: 'eu',
      GE_ENGINE: 'engine-test',
      GE_SURFACE_COMMANDER_SKILL: 'commander',
      GE_ACCESS_TOKEN: 'synthetic',
    });
    expect(provider?.assistant.location).toBe('eu');
    expect(await provider?.getAccessToken()).toBe('synthetic');
  });
});
