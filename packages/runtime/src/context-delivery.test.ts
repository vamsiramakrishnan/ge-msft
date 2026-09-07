import { describe, expect, it, vi } from 'vitest';
import type { ActuationRequest, AssistRequest, ResolvedContext, SseEvent } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { BRIEF_REF_ID } from './context-model.js';

const error: SseEvent = { type: 'error', code: 'provider_error', message: 'Request failed.' };
const blocked: SseEvent = { type: 'policy', verdict: 'block', reason: 'Request blocked.' };
const done: SseEvent = { type: 'done' };
const drain = async (events: AsyncGenerator<SseEvent>): Promise<void> => {
  for await (const _ of events) void _;
};
function fixture(responses: SseEvent[][], beforeEvent?: (event: SseEvent, call: number) => void) {
  const contexts: ResolvedContext[][] = [];
  const writes: ActuationRequest[] = [];
  const client = {
    async *stream(_request: AssistRequest, options: StreamOptions): AsyncGenerator<SseEvent> {
      const call = contexts.length;
      contexts.push(structuredClone(options.context ?? []));
      for (const event of responses[call] ?? []) {
        beforeEvent?.(event, call);
        yield event;
      }
    },
  } as unknown as StreamAssistClient;
  const bridge: DocBridge = {
    surface: 'word',
    getCapabilities: () => ({
      surface: 'word',
      contextKinds: [],
      actuations: [{ kind: 'tracked-change', surface: 'word', title: 'Suggest', reversible: true }],
    }),
    listContext: async () => [],
    resolveContext: async () => [],
    actuate: async (request) => {
      writes.push(request);
      return {
        ok: true,
        changeId: request.changeId,
        kind: request.kind,
        verification: { status: 'verified' },
      };
    },
  };
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
  });
  session.model.note('Required pending context');
  return { session, contexts, writes };
}
const containsNote = (entries: ResolvedContext[]): boolean =>
  entries.some(
    (entry) =>
      entry.ref.id === BRIEF_REF_ID &&
      entry.value.as === 'text' &&
      entry.value.text.includes('Required pending context'),
  );

describe.each(['fold', 'prime'] as const)('context delivery confirmation for %s', (mode) => {
  const send = async (session: AssistSession, signal?: AbortSignal): Promise<void> => {
    if (mode === 'fold') await drain(session.ask('Continue.', { signal }));
    else await session.commit('prime', { signal });
  };
  it.each([
    ['error', [error]],
    ['policy block', [blocked]],
    ['unconfirmed EOF', [{ type: 'token', text: 'Partial response' }]],
    ['error after done', [done, error]],
    ['policy after done', [done, blocked]],
  ] as const)(
    'keeps notes pending after %s and sends them on the next successful request',
    async (_name, events) => {
      const f = fixture([[...events], [done]]);
      await send(f.session);
      expect(f.session.model.hasPending).toBe(true);
      expect(containsNote(f.contexts[0]!)).toBe(true);
      if (mode === 'fold') {
        const expectedStatus =
          _name === 'unconfirmed EOF'
            ? 'incomplete'
            : _name.includes('policy')
              ? 'blocked'
              : 'failed';
        expect(f.session.executions.list().at(-1)?.status).toBe(expectedStatus);
        expect(
          f.session.context.list().find((entry) => entry.ref.id === BRIEF_REF_ID),
        ).toBeUndefined();
      }
      await send(f.session);
      expect(containsNote(f.contexts[1]!)).toBe(true);
      expect(f.session.model.hasPending).toBe(false);
      f.session.dispose();
    },
  );

  it('keeps notes pending when cancellation occurs before a nominal done event', async () => {
    const controller = new AbortController();
    const f = fixture([[done]], () => controller.abort());
    if (mode === 'fold') await expect(send(f.session, controller.signal)).rejects.toThrow();
    else await send(f.session, controller.signal).catch(() => undefined);
    expect(f.session.model.hasPending).toBe(true);
    f.session.dispose();
  });

  it('commits only the version delivered before newer notes arrive', async () => {
    const f = fixture([[done]], () => f.session.model.note('Newer context, not sent yet'));
    await send(f.session);
    const pending = f.session.model.pendingBrief()!.entries;
    expect(containsNote(pending)).toBe(false);
    expect(pending[0]!.value).toMatchObject({
      text: expect.stringContaining('Newer context, not sent yet'),
    });
    f.session.dispose();
  });
});

describe('provider completion gates command parsing', () => {
  const program = '```cmd\nsuggest "a" => "b"\nfinish when=verified\n```';
  it.each(['partial prose', 'closed command fence'] as const)(
    'rejects %s at EOF without approval, execution or a completed ledger record',
    async (shape) => {
      const f = fixture([
        [
          {
            type: 'token',
            text: shape === 'closed command fence' ? program : 'I will inspect it.',
          },
        ],
      ]);
      const approvePlan = vi.fn(() => true);
      const events = [];
      for await (const event of f.session.runCommands('Apply the requested change.', {
        approvePlan,
      }))
        events.push(event);
      expect(events).toContainEqual(
        expect.objectContaining({ type: 'error', code: 'incomplete_response' }),
      );
      expect(
        events.some(
          (event) =>
            event.type === 'command' || event.type === 'plan-preview' || event.type === 'done',
        ),
      ).toBe(false);
      expect(approvePlan).not.toHaveBeenCalled();
      expect(f.writes).toEqual([]);
      expect(f.session.executions.list().at(-1)).toMatchObject({
        status: 'incomplete',
        modelTurns: 1,
        effects: [],
      });
      expect(f.session.model.hasPending).toBe(true);
      f.session.dispose();
    },
  );

  it('allows the identical command only after its provider completion event', async () => {
    const f = fixture([[{ type: 'token', text: program }, done]]);
    const approvePlan = vi.fn(() => true);
    const events = [];
    for await (const event of f.session.runCommands('Apply the requested change.', { approvePlan }))
      events.push(event);
    expect(approvePlan).toHaveBeenCalledOnce();
    expect(f.writes).toHaveLength(1);
    expect(events.some((event) => event.type === 'done' && 'turn' in event)).toBe(true);
    expect(f.session.executions.list().at(-1)?.status).toBe('completed');
    f.session.dispose();
  });
});
