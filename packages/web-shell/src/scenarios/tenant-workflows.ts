import {
  TenantWorkflowReportSchema,
  type TenantWorkflowReport,
  type TenantWorkflowAttempt,
} from '@ge/contracts';
import { StreamAssistClient, parseJsonArrayStream } from '@ge/gemini-client';
import type { SseEvent } from '@ge/contracts';
import type { CommandLoopEvent } from '@ge/runtime';
import {
  TENANT_FLAGSHIP_MATRIX,
  tenantWorkflowFixture,
  tenantWorkflowSafetyCanaries,
} from './tenant-workflow-fixture.js';
import type { TenantWorkflowProvider } from './tenant-workflow-env.js';

export { TENANT_FLAGSHIP_MATRIX } from './tenant-workflow-fixture.js';
export interface TenantWorkflowSuiteOptions {
  commit: string;
  dirty: boolean;
  repetitions?: number;
  providerScope?: 'live' | 'replay';
  provider?: TenantWorkflowProvider;
  /** Across the complete matrix, including any provider retry. */
  maxModelCalls?: number;
  signal?: AbortSignal;
}
let running = false;
const record = (value: unknown): Record<string, unknown> | undefined =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : undefined;
function requestedRoute(raw: RequestInit | undefined, expected: string): boolean {
  try {
    const body = record(JSON.parse(String(raw?.body)));
    const agents = record(body?.agentsSpec)?.agentSpecs;
    return (
      body?.isSessionLess === true &&
      !body.session &&
      Array.isArray(agents) &&
      agents.some((agent) => record(agent)?.agentId === expected)
    );
  } catch {
    return false;
  }
}
async function observedInvocation(
  response: Response,
  expected: readonly string[],
  signal: AbortSignal,
): Promise<boolean> {
  if (!response.ok || !response.body) return false;
  let bytes = 0;
  const bounded = response.body.pipeThrough(
    new TransformStream<Uint8Array, Uint8Array>({
      transform(chunk, controller) {
        bytes += chunk.byteLength;
        if (bytes > 1024 * 1024) throw new Error('Observation budget reached.');
        controller.enqueue(chunk);
      },
    }),
    { signal },
  );
  let observed = false;
  try {
    for await (const chunk of parseJsonArrayStream(bounded)) {
      const skills = record(chunk)?.invokedSkills;
      if (!Array.isArray(skills) || skills.length > 64) continue;
      for (const skill of skills) {
        const value = record(skill);
        const name = typeof value?.name === 'string' ? value.name : value?.displayName;
        if (typeof name === 'string' && name.length <= 1024 && expected.includes(name))
          observed = true;
      }
    }
    return observed;
  } catch {
    return false;
  }
}
function unusedClient(): StreamAssistClient {
  return new StreamAssistClient(
    {
      getAccessToken: async () => {
        throw new Error('A deterministic fixture requested a model.');
      },
    },
    {
      assistant: { project: 'synthetic-test', location: 'eu', engine: 'synthetic-test' },
      identity: 'synthetic-acceptance',
    },
    async () => {
      throw new Error('A deterministic fixture requested a provider.');
    },
  );
}
function quantile(values: number[], p: number): number {
  if (!values.length) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  return sorted[Math.max(0, Math.ceil(sorted.length * p) - 1)]!;
}
/** Runs live-model or replayed provider calls over actual runtime + labelled simulated Office hosts. */
export async function runFlagshipTenantSuite(
  options: TenantWorkflowSuiteOptions,
): Promise<TenantWorkflowReport> {
  if (running) throw new Error('A flagship suite already owns the simulated Office globals.');
  if (!/^[a-f0-9]{40}$/.test(options.commit)) throw new Error('A full commit SHA is required.');
  const repetitions = options.repetitions ?? 1;
  const budget = options.maxModelCalls ?? 12;
  if (!Number.isInteger(repetitions) || repetitions < 1 || repetitions > 5)
    throw new Error('Use one to five repetitions.');
  if (!Number.isInteger(budget) || budget < 1 || budget > 24)
    throw new Error('Use a model-call budget of one to 24.');
  running = true;
  const started = performance.now();
  const attempts: TenantWorkflowAttempt[] = [];
  let totalCalls = 0;
  let providerUnavailable = false;
  let safety = { sourceFreshnessVerified: false, approvalRefusalVerified: false };
  try {
    options.signal?.throwIfAborted();
    try {
      safety = await tenantWorkflowSafetyCanaries(unusedClient());
    } catch {
      /* Absent safety proof cannot pass release acceptance. */
    }
    for (let repetition = 1; repetition <= repetitions; repetition++) {
      for (const spec of TENANT_FLAGSHIP_MATRIX) {
        const attempt: TenantWorkflowAttempt = {
          workflowId: spec.id,
          surface: spec.surface,
          repetition,
          status: 'unavailable',
          modelRequired: spec.modelRequired,
          modelCalls: 0,
          requestRoutingVerified: false,
          skillInvocationObserved: false,
          outcomeVerified: false,
          durationMs: 0,
          issueCodes: [],
        };
        attempts.push(attempt);
        if (options.signal?.aborted) {
          attempt.issueCodes.push('cancelled');
          continue;
        }
        if (
          spec.modelRequired &&
          (!options.provider ||
            providerUnavailable ||
            totalCalls >= budget ||
            (options.dirty && options.providerScope !== 'replay'))
        ) {
          attempt.issueCodes.push(
            !options.provider
              ? 'configuration_missing'
              : options.dirty && options.providerScope !== 'replay'
                ? 'dirty_worktree'
                : totalCalls >= budget
                  ? 'query_budget'
                  : 'provider_unavailable',
          );
          continue;
        }
        const begin = performance.now();
        const controller = new AbortController();
        const abort = () => controller.abort();
        options.signal?.addEventListener('abort', abort, { once: true });
        const timer = setTimeout(abort, 180000);
        const observations: Promise<boolean>[] = [];
        const requested: boolean[] = [];
        let observationBudget = false;
        const provider = options.provider;
        const expected = provider
          ? [provider.commander.resource, provider.commander.resource.split('/').at(-1)!]
          : [];
        const fetchImpl: typeof fetch = async (input, init) => {
          if (totalCalls >= budget) {
            observationBudget = true;
            throw new Error('Workflow model-call budget reached.');
          }
          totalCalls++;
          attempt.modelCalls++;
          requested.push(requestedRoute(init, provider!.commander.resource.split('/').at(-1)!));
          const response = await (provider!.fetchImpl ?? fetch)(input, init);
          if (!response.ok) providerUnavailable = true;
          observations.push(observedInvocation(response.clone(), expected, controller.signal));
          return response;
        };
        const client = spec.modelRequired
          ? new StreamAssistClient(
              { getAccessToken: () => provider!.getAccessToken() },
              {
                assistant: provider!.assistant,
                identity: 'synthetic-flagship-acceptance',
                commandSkills: [provider!.commander.resource],
                commandSkillMentions: [
                  {
                    label: provider!.commander.label,
                    uri: provider!.commander.resource.split('/').at(-1)!,
                  },
                ],
              },
              fetchImpl,
              { maxAttempts: 1 },
            )
          : unusedClient();
        let fixture: Awaited<ReturnType<typeof tenantWorkflowFixture>> | undefined;
        try {
          fixture = await tenantWorkflowFixture(spec.id, client);
          const events: Array<SseEvent | CommandLoopEvent> = [];
          for await (const event of fixture.run({
            approvePlan: () => true,
            signal: controller.signal,
            maxTurns: 3,
            maxCommandsPerTurn: 24,
            maxWritesPerTurn: 4,
          })) {
            if (events.length >= 512) throw new Error('Workflow event budget reached.');
            events.push(event);
          }
          const invoked = await Promise.all(observations);
          if (
            spec.modelRequired &&
            (!attempt.modelCalls ||
              events.some(
                (event) =>
                  event.type === 'error' &&
                  (/^http_/.test(event.code) ||
                    ['network', 'auth_exchange', 'timeout', 'incomplete_response'].includes(
                      event.code,
                    )),
              ))
          )
            providerUnavailable = true;
          attempt.requestRoutingVerified = requested.length > 0 && requested.every(Boolean);
          attempt.skillInvocationObserved = invoked.some(Boolean);
          attempt.outcomeVerified = fixture.outcome(events);
          if (controller.signal.aborted) attempt.issueCodes.push('cancelled');
          else if (observationBudget) attempt.issueCodes.push('query_budget');
          else if (spec.modelRequired && providerUnavailable)
            attempt.issueCodes.push('provider_unavailable');
          else if (!attempt.outcomeVerified) attempt.issueCodes.push('outcome_mismatch');
          if (spec.modelRequired && !attempt.requestRoutingVerified)
            attempt.issueCodes.push('routing_unobserved');
          if (spec.modelRequired && !attempt.skillInvocationObserved)
            attempt.issueCodes.push('skill_invocation_unobserved');
          attempt.status = attempt.issueCodes.includes('outcome_mismatch')
            ? 'failed'
            : attempt.issueCodes.length
              ? 'unavailable'
              : 'passed';
        } catch {
          attempt.issueCodes.push(
            controller.signal.aborted
              ? 'cancelled'
              : observationBudget
                ? 'query_budget'
                : 'unexpected_error',
          );
          attempt.status = spec.modelRequired ? 'unavailable' : 'failed';
        } finally {
          controller.abort();
          clearTimeout(timer);
          options.signal?.removeEventListener('abort', abort);
          fixture?.dispose();
          await Promise.allSettled(observations);
          attempt.durationMs = performance.now() - begin;
        }
      }
    }
    const passed = attempts.filter((attempt) => attempt.status === 'passed').length;
    const failed = attempts.filter((attempt) => attempt.status === 'failed').length;
    const unavailable = attempts.length - passed - failed;
    const durations = attempts
      .filter((attempt) => attempt.durationMs > 0)
      .map((attempt) => attempt.durationMs);
    return TenantWorkflowReportSchema.parse({
      version: 1,
      commit: options.commit,
      dirty: options.dirty,
      providerScope: options.provider ? (options.providerScope ?? 'live') : 'unavailable',
      transport: 'public',
      hostScope: 'simulated',
      identityScope: !options.provider
        ? 'unavailable'
        : options.providerScope === 'replay'
          ? 'synthetic-replay'
          : 'explicit-developer-token',
      status:
        failed || !safety.approvalRefusalVerified || !safety.sourceFreshnessVerified
          ? 'FAIL'
          : unavailable || options.dirty
            ? 'BLOCKED_EXTERNAL'
            : 'PASSED',
      ...safety,
      attempts,
      summary: {
        requested: attempts.length,
        passed,
        failed,
        unavailable,
        modelCalls: totalCalls,
        durationMs: performance.now() - started,
        p50Ms: quantile(durations, 0.5),
        p95Ms: quantile(durations, 0.95),
      },
    });
  } finally {
    running = false;
  }
}
