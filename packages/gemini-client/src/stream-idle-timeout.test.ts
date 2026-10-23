import { describe, it, expect, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import { StreamAssistClient } from './stream-assist.js';

const cfg = { assistant: { project: 'proj', location: 'eu', engine: 'eng1' }, identity: 'u@x' };
const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };
const req: AssistRequest = {
  intent: 'ask',
  query: 'q',
  unit: { connectors: [], surfaceContext: { kind: 'word', selection: 's' } },
};
const frame = (text: string): string =>
  JSON.stringify({
    answer: { state: 'IN_PROGRESS', replies: [{ groundedContent: { content: { text } } }] },
  });
const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** A body that emits `pieces` with `gapMs` between them, then optionally stalls forever. */
function timedBody(pieces: string[], gapMs: number, stall = false): ReadableStream<Uint8Array> {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream({
    async pull(controller) {
      if (i > 0) await sleep(gapMs);
      if (i < pieces.length) controller.enqueue(enc.encode(pieces[i++]!));
      else if (stall) await new Promise(() => {});
      else controller.close();
    },
  });
}

async function collect(gen: AsyncGenerator<SseEvent>, perEventMs = 0): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const ev of gen) {
    out.push(ev);
    if (perEventMs) await sleep(perEventMs);
  }
  return out;
}

describe('streamAssist idle timeout', () => {
  it('times out when the engine never answers the request', async () => {
    const fetch = vi.fn(
      (_url: string, init?: RequestInit) =>
        new Promise<Response>((_, reject) => {
          // Like real fetch: an already-aborted signal rejects at once.
          if (init?.signal?.aborted) reject(new DOMException('x', 'AbortError'));
          init?.signal?.addEventListener('abort', () =>
            reject(new DOMException('x', 'AbortError')),
          );
        }),
    );
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetch as unknown as typeof globalThis.fetch,
      {},
    );
    const events = await collect(client.stream(req, { idleTimeoutMs: 40 }));
    expect(events.at(-1)).toMatchObject({ type: 'error', code: 'timeout' });
  });

  it('times out when the stream stalls mid-answer, keeping what already arrived', async () => {
    const body = timedBody(['[', frame('Hello'), ','], 5, true);
    const fetch = vi.fn(async () => new Response(body, { status: 200 }));
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetch as unknown as typeof globalThis.fetch,
      {},
    );
    const events = await collect(client.stream(req, { idleTimeoutMs: 40 }));
    expect(events.some((e) => e.type === 'token')).toBe(true);
    expect(events.at(-1)).toMatchObject({ type: 'error', code: 'timeout' });
  });

  it('does not time out a slow stream that keeps sending', async () => {
    const body = timedBody(['[', frame('a'), ',', frame('b'), ']'], 25);
    const fetch = vi.fn(async () => new Response(body, { status: 200 }));
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetch as unknown as typeof globalThis.fetch,
      {},
    );
    const events = await collect(client.stream(req, { idleTimeoutMs: 40 }));
    expect(events.some((e) => e.type === 'error')).toBe(false);
    expect(events.at(-1)).toMatchObject({ type: 'done' });
  });

  it('does not count time the consumer spends between reads', async () => {
    const body = timedBody(['[', frame('a'), ',', frame('b'), ']'], 0);
    const fetch = vi.fn(async () => new Response(body, { status: 200 }));
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetch as unknown as typeof globalThis.fetch,
      {},
    );
    const events = await collect(client.stream(req, { idleTimeoutMs: 30 }), 60);
    expect(events.some((e) => e.type === 'error')).toBe(false);
    expect(events.at(-1)).toMatchObject({ type: 'done' });
  });

  it('a user cancel still cancels (not reported as a timeout)', async () => {
    const abort = new AbortController();
    const fetch = vi.fn(
      (_url: string, init?: RequestInit) =>
        new Promise<Response>((_, reject) => {
          // Like real fetch: an already-aborted signal rejects at once.
          if (init?.signal?.aborted) reject(new DOMException('x', 'AbortError'));
          init?.signal?.addEventListener('abort', () =>
            reject(new DOMException('x', 'AbortError')),
          );
        }),
    );
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetch as unknown as typeof globalThis.fetch,
      {},
    );
    setTimeout(() => abort.abort(), 10);
    const events = await collect(client.stream(req, { idleTimeoutMs: 1000, signal: abort.signal }));
    expect(events.some((e) => e.type === 'error' && e.code === 'timeout')).toBe(false);
  });
});
