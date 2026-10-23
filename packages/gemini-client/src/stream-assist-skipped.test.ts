import { describe, it, expect, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import { SKIP_RETRY_FRAME, StreamAssistClient } from './stream-assist.js';

/**
 * streamAssist answers a message it judges not answer-seeking ("hi", "thanks") with state SKIPPED
 * and no text (seen live 2026-10-15: `NON_ASSIST_SEEKING_QUERY_IGNORED`). The client must not end
 * such a turn in silence, and must never let a policy skip through.
 */

const cfg = {
  assistant: { project: 'proj', location: 'global', engine: 'eng1' },
  identity: 'u@acme',
};
const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };
const noSleep = { sleep: async () => {}, random: () => 0 };
const CHAT = { retrySkipped: true };

function req(query = 'hii'): AssistRequest {
  return { intent: 'ask', query, unit: { connectors: [], surfaceContext: { kind: 'excel' } } };
}

const skipped = (reasons: string[] = ['NON_ASSIST_SEEKING_QUERY_IGNORED'], extra = {}) => [
  { ...extra, answer: { state: 'SKIPPED', assistSkippedReasons: reasons } },
];
const answered = (text: string) => [
  { answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] } },
];

/** A fetch that serves the given response bodies in order and records each request body. */
function sequence(...bodies: unknown[][]) {
  const sent: Array<{ query: { text?: string }; session?: string }> = [];
  const f = vi.fn(async (_url: string, init: { body: string }) => {
    sent.push(JSON.parse(init.body));
    const body = bodies[Math.min(sent.length - 1, bodies.length - 1)]!;
    return new Response(JSON.stringify(body), { status: 200 });
  });
  return { f, sent, queries: () => sent.map((b) => b.query.text ?? '') };
}

async function collect(gen: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const ev of gen) out.push(ev);
  return out;
}

const text = (events: SseEvent[]): string =>
  events.map((e) => (e.type === 'token' ? e.text : '')).join('');

describe('a skipped streamAssist answer', () => {
  it('on a chat question: asked again once, framed, and the second answer is what the pane gets', async () => {
    const { f, queries } = sequence(skipped(), answered('Hello! How can I help?'));
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('hii'), CHAT));
    expect(f).toHaveBeenCalledTimes(2);
    expect(queries()[0]).toContain('hii');
    expect(queries()[0]).not.toContain(SKIP_RETRY_FRAME);
    expect(queries()[1]).toContain(`${SKIP_RETRY_FRAME}hii`);
    expect(text(events)).toBe('Hello! How can I help?');
    expect(events.filter((e) => e.type === 'done')).toHaveLength(1);
    expect(events.some((e) => e.type === 'error')).toBe(false);
  });

  it('keeps the retry in the session the skipped call opened', async () => {
    const { f, sent } = sequence(
      skipped(['NON_ASSIST_SEEKING_QUERY_IGNORED'], {
        sessionInfo: { session: 'projects/p/sessions/s1' },
      }),
      answered('Hi!'),
    );
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    await collect(client.stream(req('hii'), CHAT));
    expect(sent[1]?.session).toBe('projects/p/sessions/s1');
  });

  it('is reported, not silent, when the second try is skipped too', async () => {
    const { f } = sequence(skipped(), skipped());
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('ok'), CHAT));
    expect(f).toHaveBeenCalledTimes(2);
    expect(events.at(-1)).toMatchObject({ type: 'error', code: 'assist_skipped' });
  });

  it('is never reframed on the planner or command route: reported at once', async () => {
    const { f } = sequence(skipped(), answered('should not be asked'));
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('protocol + <doc_state> + REQUEST: hii')));
    expect(f).toHaveBeenCalledTimes(1);
    expect(events.at(-1)).toMatchObject({ type: 'error', code: 'assist_skipped' });
  });

  it('is not asked again for an unknown or unspecified skip reason', async () => {
    for (const reasons of [[], ['ASSIST_SKIPPED_REASON_UNSPECIFIED'], ['SOME_FUTURE_REASON']]) {
      const { f } = sequence(skipped(reasons), answered('should not be asked'));
      const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
      const events = await collect(client.stream(req('hii'), CHAT));
      expect(f).toHaveBeenCalledTimes(1);
      expect(events.at(-1)).toMatchObject({ type: 'error', code: 'assist_skipped' });
    }
  });

  it('drops citations from the skipped call; the answer keeps only its own', async () => {
    const citeOnly = [
      {
        answer: {
          state: 'SKIPPED',
          assistSkippedReasons: ['NON_ASSIST_SEEKING_QUERY_IGNORED'],
          replies: [
            {
              groundedContent: {
                textGroundingMetadata: {
                  references: [{ documentMetadata: { title: 'Stale source', uri: 'https://x/1' } }],
                },
              },
            },
          ],
        },
      },
    ];
    const { f } = sequence(citeOnly, answered('Hi!'));
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('hii'), CHAT));
    expect(events.some((e) => e.type === 'citation')).toBe(false);
    expect(text(events)).toBe('Hi!');
  });
});

describe('a policy skip', () => {
  it('is a block when the engine skipped the query for a customer policy', async () => {
    const { f } = sequence(skipped(['CUSTOMER_POLICY_VIOLATION']));
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('something blocked'), CHAT));
    expect(f).toHaveBeenCalledTimes(1);
    expect(events.map((e) => e.type)).toEqual(['policy', 'done']);
  });

  it('blocks even after text streamed: no provenance, nothing usable', async () => {
    const body = [
      {
        answer: {
          state: 'IN_PROGRESS',
          replies: [{ groundedContent: { content: { text: '```cmd\nset A1 1\n```' } } }],
        },
      },
      { answer: { state: 'SKIPPED', assistSkippedReasons: ['CUSTOMER_POLICY_VIOLATION'] } },
    ];
    const { f } = sequence(body);
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('do something'), CHAT));
    expect(events.some((e) => e.type === 'policy')).toBe(true);
    expect(events.some((e) => e.type === 'provenance')).toBe(false);
    expect(events.at(-1)).toMatchObject({ type: 'done' });
  });
});

describe('a normal answer', () => {
  it('is left alone: one call, no framing', async () => {
    const { f, queries } = sequence(answered('The SLA is 99.5%.'));
    const client = new StreamAssistClient(tokens, cfg, f as never, noSleep);
    const events = await collect(client.stream(req('What is the SLA?'), CHAT));
    expect(f).toHaveBeenCalledTimes(1);
    expect(queries()[0]).not.toContain(SKIP_RETRY_FRAME);
    expect(text(events)).toBe('The SLA is 99.5%.');
  });
});
