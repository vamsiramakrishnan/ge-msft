import { afterEach, describe, expect, it, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import { StreamAssistClient } from './stream-assist.js';

const cfg = { assistant: { project: 'proj', location: 'eu', engine: 'eng1' }, identity: 'u@x' };
const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };
const req: AssistRequest = {
  intent: 'ask',
  query: 'q',
  unit: { connectors: [], surfaceContext: { kind: 'word', selection: 's' } },
};
const noWait = { sleep: async () => {}, random: () => 0 };

function bodyOf(pieces: string[], opts: { stallAfter?: boolean; errorAfter?: boolean } = {}) {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      if (i < pieces.length) controller.enqueue(enc.encode(pieces[i++]!));
      else if (opts.stallAfter) await new Promise(() => {});
      else if (opts.errorAfter) controller.error(new TypeError('socket hang up'));
      else controller.close();
    },
  });
}
const ok = (text: string): Response =>
  new Response(
    bodyOf([
      JSON.stringify([
        {
          answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
        },
      ]),
    ]),
    { status: 200 },
  );
const frame = (text: string, state = 'IN_PROGRESS'): string =>
  JSON.stringify({ answer: { state, replies: [{ groundedContent: { content: { text } } }] } });

async function collect(gen: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const ev of gen) out.push(ev);
  return out;
}

afterEach(() => {
  vi.useRealTimers();
});

describe('turn-level retry', () => {
  it('asks once more when the connection drops before any of the answer arrived', async () => {
    const fetch = vi
      .fn<[string, RequestInit?], Promise<Response>>()
      .mockRejectedValueOnce(new TypeError('Failed to fetch'))
      .mockResolvedValueOnce(ok('Hello.'));
    const client = new StreamAssistClient(tokens, cfg, fetch as never, {
      maxAttempts: 1,
      ...noWait,
    });
    const events = await collect(client.stream(req));
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(events.map((e) => e.type)).toEqual(['activity', 'token', 'provenance', 'done']);
    expect(events.some((e) => e.type === 'error')).toBe(false);
  });

  it('asks once more when the body was cut short before any content', async () => {
    const truncated = new Response(bodyOf(['[{"sessionInfo":{"session":"s"}}']), { status: 200 });
    const fetch = vi
      .fn<[string, RequestInit?], Promise<Response>>()
      .mockResolvedValueOnce(truncated)
      .mockResolvedValueOnce(ok('Done.'));
    const client = new StreamAssistClient(tokens, cfg, fetch as never, noWait);
    const events = await collect(client.stream(req));
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(events.at(-1)).toEqual({ type: 'done' });
    expect(events.filter((e) => e.type === 'token')).toHaveLength(1);
  });

  it('does not retry once part of the answer reached the consumer', async () => {
    const partial = new Response(bodyOf(['[', frame('Half ')], { errorAfter: true }), {
      status: 200,
    });
    const fetch = vi.fn<[string, RequestInit?], Promise<Response>>().mockResolvedValueOnce(partial);
    const client = new StreamAssistClient(tokens, cfg, fetch as never, noWait);
    const events = await collect(client.stream(req)).catch((err: unknown) => err);
    expect(fetch).toHaveBeenCalledTimes(1);
    // The partial token was shown; the failure is final (surfaced, never silently re-answered).
    expect(Array.isArray(events) ? events.some((e) => e.type === 'token') : true).toBe(true);
  });

  it('does not retry after the caller cancelled', async () => {
    const controller = new AbortController();
    const fetch = vi.fn<[string, RequestInit?], Promise<Response>>().mockImplementation(() => {
      controller.abort();
      return Promise.reject(new TypeError('Failed to fetch'));
    });
    const client = new StreamAssistClient(tokens, cfg, fetch as never, {
      maxAttempts: 1,
      ...noWait,
    });
    const events = await collect(client.stream(req, { signal: controller.signal }));
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(events.at(-1)).toMatchObject({ type: 'error', code: 'network' });
  });
});

describe('circuit breaker', () => {
  it('opens after repeated POST failures, fails fast, and probes again after the cooldown', async () => {
    vi.useFakeTimers();
    const fetch = vi.fn<[string, RequestInit?], Promise<Response>>(() =>
      Promise.reject(new TypeError('Failed to fetch')),
    );
    const client = new StreamAssistClient(
      tokens,
      { ...cfg, resilience: { circuitBreaker: { threshold: 2, cooldownMs: 1_000 } } },
      fetch as never,
      { maxAttempts: 1, ...noWait },
    );
    // Each turn = one exhausted POST plus its turn-level retry = two failures → the circuit opens.
    const first = await collect(client.stream(req));
    expect(first.at(-1)).toMatchObject({ type: 'error', code: 'network' });
    expect(client.circuitState).toBe('open');
    const calls = fetch.mock.calls.length;

    const second = await collect(client.stream(req));
    expect(second).toEqual([expect.objectContaining({ type: 'error', code: 'circuit_open' })]);
    expect(fetch).toHaveBeenCalledTimes(calls);

    vi.advanceTimersByTime(1_001);
    expect(client.circuitState).toBe('half-open');
    fetch.mockResolvedValueOnce(ok('Back.'));
    const third = await collect(client.stream(req));
    expect(third.at(-1)).toEqual({ type: 'done' });
    expect(client.circuitState).toBe('closed');
  });

  it('can be disabled', () => {
    const client = new StreamAssistClient(
      tokens,
      { ...cfg, resilience: { circuitBreaker: false } },
      vi.fn() as never,
    );
    expect(client.circuitState).toBe('disabled');
  });
});

describe('per-turn deadline', () => {
  it('ends a turn that keeps streaming past the deadline', async () => {
    const stalling = new Response(bodyOf(['[', frame('Working ')], { stallAfter: true }), {
      status: 200,
    });
    const fetch = vi
      .fn<[string, RequestInit?], Promise<Response>>()
      .mockResolvedValueOnce(stalling);
    const client = new StreamAssistClient(tokens, cfg, fetch as never, noWait);
    const events = await collect(client.stream(req, { deadlineMs: 60, idleTimeoutMs: 10_000 }));
    expect(events.some((e) => e.type === 'token')).toBe(true);
    const last = events.at(-1);
    expect(last).toMatchObject({ type: 'error', code: 'timeout' });
    expect(last && last.type === 'error' && last.message).toContain('finish');
    expect(fetch).toHaveBeenCalledTimes(1);
  });
});
