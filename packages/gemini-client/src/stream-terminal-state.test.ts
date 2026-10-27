import { describe, it, expect, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import { StreamAssistClient } from './stream-assist.js';
import type { GeminiClientConfig } from './config.js';
import { parseJsonArrayStream, type JsonStreamStatus } from './json-stream.js';

/**
 * `done` is a claim that the provider finished the turn. A clean EOF is not that claim: a proxy
 * or the network can close the body mid-array and everything received so far still parses.
 * These tests pin the rule — `done` only after a terminal answer state or the array's closing
 * bracket; otherwise `incomplete_response`, which the runtime treats as "commands not executed".
 */

const ASSISTANT = { project: 'proj', location: 'eu', engine: 'eng1' };

function cfg(): GeminiClientConfig {
  return { assistant: ASSISTANT, identity: 'v.k@acme' };
}

function assistReq(): AssistRequest {
  return {
    intent: 'ask',
    query: 'What is the SLA?',
    unit: {
      connectors: [],
      surfaceContext: { kind: 'word', selection: 'available 99.5% of time' },
    },
  };
}

function streamOf(pieces: string[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream({
    pull(controller) {
      if (i < pieces.length) controller.enqueue(enc.encode(pieces[i++]!));
      else controller.close();
    },
  });
}

const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };
const noSleep = { sleep: async () => {}, random: () => 0 };

async function collect(gen: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const ev of gen) out.push(ev);
  return out;
}

function rawFetch(pieces: string[]) {
  return vi.fn(async () => new Response(streamOf(pieces), { status: 200 }));
}

const reply = (text: string, state?: string) => ({
  answer: {
    ...(state ? { state } : {}),
    replies: [{ groundedContent: { content: { text } } }],
  },
});

describe('StreamAssistClient — end-of-stream truth', () => {
  it('a body truncated mid-object yields incomplete_response and no done', async () => {
    const full = JSON.stringify([reply('Hello', 'IN_PROGRESS'), reply(' world', 'SUCCEEDED')]);
    // Cut inside the second object, before any terminal state is visible.
    const cut = full.indexOf('SUCCEEDED') - 3;
    const client = new StreamAssistClient(
      tokens,
      cfg(),
      rawFetch([full.slice(0, cut)]) as never,
      noSleep,
    );
    const events = await collect(client.stream(assistReq()));
    expect(events.some((e) => e.type === 'done')).toBe(false);
    expect(events.some((e) => e.type === 'provenance')).toBe(false);
    const last = events[events.length - 1]!;
    expect(last).toMatchObject({ type: 'error', code: 'incomplete_response' });
    // The text that did arrive still streamed; the runtime decides what to do with it.
    expect(events.some((e) => e.type === 'token' && e.text === 'Hello')).toBe(true);
  });

  it('a body that ends after a complete object but before the closing bracket is incomplete', async () => {
    const full = JSON.stringify([reply('Hello', 'IN_PROGRESS'), reply(' world', 'IN_PROGRESS')]);
    const client = new StreamAssistClient(
      tokens,
      cfg(),
      rawFetch([full.slice(0, full.length - 1)]) as never,
      noSleep,
    );
    const events = await collect(client.stream(assistReq()));
    expect(events.some((e) => e.type === 'done')).toBe(false);
    expect(events[events.length - 1]).toMatchObject({ type: 'error', code: 'incomplete_response' });
  });

  it('a closed array without a SUCCEEDED state still completes with done', async () => {
    const body = JSON.stringify([reply('Hi.')]);
    const client = new StreamAssistClient(tokens, cfg(), rawFetch([body]) as never, noSleep);
    const events = await collect(client.stream(assistReq()));
    expect(events.filter((e) => e.type === 'done')).toHaveLength(1);
    expect(events.some((e) => e.type === 'error')).toBe(false);
    expect(events[events.length - 1]!.type).toBe('done');
  });

  it('a SUCCEEDED state with a closed array emits done exactly once', async () => {
    const body = JSON.stringify([reply('Hi.', 'SUCCEEDED')]);
    const client = new StreamAssistClient(tokens, cfg(), rawFetch([body]) as never, noSleep);
    const events = await collect(client.stream(assistReq()));
    expect(events.filter((e) => e.type === 'done')).toHaveLength(1);
    expect(events[events.length - 1]!.type).toBe('done');
  });

  it('a SUCCEEDED frame counts as terminal even when the closing bracket never arrives', async () => {
    const full = JSON.stringify([reply('Hi.', 'SUCCEEDED')]);
    const client = new StreamAssistClient(
      tokens,
      cfg(),
      rawFetch([full.slice(0, full.length - 1)]) as never,
      noSleep,
    );
    const events = await collect(client.stream(assistReq()));
    expect(events.filter((e) => e.type === 'done')).toHaveLength(1);
  });
});

describe('parseJsonArrayStream — framing status', () => {
  async function run(pieces: string[]): Promise<{ objs: unknown[]; status: JsonStreamStatus }> {
    const status: JsonStreamStatus = { closed: false, incomplete: false };
    const objs: unknown[] = [];
    for await (const o of parseJsonArrayStream(streamOf(pieces), status)) objs.push(o);
    return { objs, status };
  }

  it('reports closed only when the top-level bracket is seen outside strings and objects', async () => {
    const { objs, status } = await run(['[{"a":"]"},', '{"b":[1,2]}', ']']);
    expect(objs).toEqual([{ a: ']' }, { b: [1, 2] }]);
    expect(status).toEqual({ closed: true, incomplete: false });
  });

  it('reports incomplete when the body ends inside an object', async () => {
    const { objs, status } = await run(['[{"a":1},{"b":']);
    expect(objs).toEqual([{ a: 1 }]);
    expect(status.closed).toBe(false);
    expect(status.incomplete).toBe(true);
  });

  it('reports neither closed nor incomplete for a clean cut between objects', async () => {
    const { status } = await run(['[{"a":1},']);
    expect(status).toEqual({ closed: false, incomplete: false });
  });
});
