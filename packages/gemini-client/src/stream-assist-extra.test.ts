import { describe, it, expect, vi } from 'vitest';
import type { AssistRequest, ResolvedContext, SseEvent } from '@ge/contracts';
import { StreamAssistClient } from './stream-assist.js';
import type { GeminiClientConfig } from './config.js';

const ASSISTANT = { project: 'proj', location: 'eu', engine: 'eng1' };

function cfg(overrides: Partial<GeminiClientConfig> = {}): GeminiClientConfig {
  return { assistant: ASSISTANT, identity: 'v.k@acme', ...overrides };
}

function assistReq(query = 'What is the SLA?'): AssistRequest {
  return {
    intent: 'assist',
    query,
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

function streamFetch(chunks: unknown[], status = 200) {
  return vi.fn(async () => new Response(streamOf([JSON.stringify(chunks)]), { status }));
}

describe('StreamAssistClient — network and failure paths', () => {
  it('yields a network error event when the POST throws (no Response)', async () => {
    const f = vi.fn(async () => {
      throw new TypeError('Failed to fetch');
    });
    const client = new StreamAssistClient(tokens, cfg(), f as never, noSleep);
    const events = await collect(client.stream(assistReq()));
    expect(events).toHaveLength(1);
    expect(events[0]).toMatchObject({
      type: 'error',
      code: 'network',
      message: 'Failed to fetch',
    });
  });

  it('maps an exhausted transient 503 to an http_503 error event', async () => {
    const f = vi.fn(async () => new Response('unavailable', { status: 503 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never, {
      ...noSleep,
      maxAttempts: 2,
    });
    const events = await collect(client.stream(assistReq()));
    expect(events[0]).toMatchObject({ type: 'error', code: 'http_503' });
    // 1 initial + 1 retry.
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('surfaces an http_500 error when a 401 retry still fails transiently', async () => {
    const invalidate = vi.fn();
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(new Response('boom', { status: 500 }));
    const client = new StreamAssistClient(
      { getAccessToken: () => Promise.resolve('t'), invalidate },
      cfg(),
      f as never,
      { ...noSleep, maxAttempts: 1 },
    );
    const events = await collect(client.stream(assistReq()));
    expect(invalidate).toHaveBeenCalledOnce();
    // The re-sent request returns a transient throw → mapped to a synthetic 500 Response.
    expect(events[0]).toMatchObject({ type: 'error', code: 'http_500' });
  });

  it('emits assist_failed when the answer state is FAILED (no policy block)', async () => {
    const chunk = {
      answer: { state: 'FAILED', replies: [] },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    expect(events.some((e) => e.type === 'error' && e.code === 'assist_failed')).toBe(true);
    // A non-blocked turn still completes with provenance + done.
    expect(events.some((e) => e.type === 'provenance')).toBe(true);
    expect(events[events.length - 1]!.type).toBe('done');
  });

  it('tolerates non-conforming frames by skipping them', async () => {
    // First frame is junk metadata (fails the schema), second is a real answer.
    const junk = { keepalive: true, foo: { bar: [1, 2] } };
    const good = {
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text: 'Hi.' } } }] },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([junk, good]) as never);
    const events = await collect(client.stream(assistReq()));
    const tokenText = events
      .filter((e): e is Extract<SseEvent, { type: 'token' }> => e.type === 'token')
      .map((e) => e.text)
      .join('');
    expect(tokenText).toBe('Hi.');
  });

  it('does not stream a reply whose content is flagged as a thought', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          { groundedContent: { content: { text: 'internal reasoning', thought: true } } },
          { groundedContent: { content: { text: 'final answer' } } },
        ],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const tokenText = events
      .filter((e): e is Extract<SseEvent, { type: 'token' }> => e.type === 'token')
      .map((e) => e.text)
      .join('');
    expect(tokenText).toBe('final answer');
    expect(tokenText).not.toContain('internal reasoning');
  });
});

describe('StreamAssistClient — grounding-support source resolution', () => {
  it('resolves a support source from inline documentMetadata (no referenceIndex)', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'abcdef' },
              textGroundingMetadata: {
                references: [],
                groundingSupports: [
                  {
                    startIndex: '0',
                    endIndex: '3',
                    sources: [
                      {
                        documentMetadata: {
                          title: 'Inline Doc',
                          uri: 'https://x/inline',
                          pageIdentifier: 'p7',
                        },
                      },
                    ],
                  },
                ],
              },
            },
          },
        ],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const support = events.find(
      (e): e is Extract<SseEvent, { type: 'grounding-support' }> => e.type === 'grounding-support',
    );
    expect(support?.sources[0]).toEqual({
      title: 'Inline Doc',
      uri: 'https://x/inline',
      locator: 'p7',
    });
  });

  it('drops a support source that has neither title nor uri', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'abcdef' },
              textGroundingMetadata: {
                references: [],
                groundingSupports: [
                  { startIndex: '0', endIndex: '3', sources: [{ documentMetadata: {} }] },
                ],
              },
            },
          },
        ],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const support = events.find(
      (e): e is Extract<SseEvent, { type: 'grounding-support' }> => e.type === 'grounding-support',
    );
    // The span is still emitted, but with no usable sources.
    expect(support?.sources).toEqual([]);
  });

  it('de-duplicates support sources that resolve to the same uri', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'abcdef' },
              textGroundingMetadata: {
                references: [{ documentMetadata: { title: 'Doc', uri: 'https://x/d' } }],
                groundingSupports: [
                  {
                    startIndex: '0',
                    endIndex: '3',
                    sources: [{ referenceIndex: 0 }, { uri: 'https://x/d', title: 'Doc again' }],
                  },
                ],
              },
            },
          },
        ],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const support = events.find(
      (e): e is Extract<SseEvent, { type: 'grounding-support' }> => e.type === 'grounding-support',
    );
    expect(support?.sources).toHaveLength(1);
    expect(support?.sources[0]?.uri).toBe('https://x/d');
  });

  it('does not re-emit a grounding-support span with identical char bounds', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'abcdef' },
              textGroundingMetadata: {
                references: [{ documentMetadata: { title: 'Doc', uri: 'https://x/d' } }],
                groundingSupports: [
                  { startIndex: '0', endIndex: '3', sources: [{ referenceIndex: 0 }] },
                  { startIndex: '0', endIndex: '3', sources: [{ referenceIndex: 0 }] },
                ],
              },
            },
          },
        ],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const supports = events.filter((e) => e.type === 'grounding-support');
    expect(supports).toHaveLength(1);
  });
});

describe('StreamAssistClient — provenance and agent id', () => {
  it('stamps provenance with a content hash of the accumulated answer and invoked skills in the agent id', async () => {
    const chunk = {
      sessionInfo: { session: 'sess_77' },
      invokedSkills: [{ displayName: 'Researcher' }, { name: 'summarize' }],
      answer: {
        state: 'SUCCEEDED',
        replies: [{ groundedContent: { content: { text: 'Grounded answer.' } } }],
      },
    };
    const client = new StreamAssistClient(tokens, cfg(), streamFetch([chunk]) as never);
    const events = await collect(client.stream(assistReq()));
    const prov = events.find(
      (e): e is Extract<SseEvent, { type: 'provenance' }> => e.type === 'provenance',
    );
    expect(prov?.payload.contentHash).toMatch(/^sha256:[0-9a-f]{64}$/);
    expect(prov?.payload.agentId).toContain('Researcher+summarize');
    expect(prov?.payload.identity).toBe('v.k@acme');
    expect(prov?.payload.sessionId).toBe('sess_77');
  });

  it("defaults identity to 'unknown' when the config omits it", async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [{ groundedContent: { content: { text: 'x' } } }],
      },
    };
    const client = new StreamAssistClient(
      tokens,
      cfg({ identity: undefined }),
      streamFetch([chunk]) as never,
    );
    const events = await collect(client.stream(assistReq()));
    const prov = events.find(
      (e): e is Extract<SseEvent, { type: 'provenance' }> => e.type === 'provenance',
    );
    expect(prov?.payload.identity).toBe('unknown');
  });
});

describe('StreamAssistClient — multipart context query', () => {
  it('sends attached context as data parts with the question last, never as instructions', async () => {
    const f = streamFetch([
      {
        answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text: 'ok' } } }] },
      },
    ]);
    const context: ResolvedContext[] = [
      {
        ref: { id: 's1', kind: 'selection', surface: 'word', title: 'Selection' },
        value: { as: 'text', text: 'attached selection' },
      },
    ];
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    await collect(client.stream(assistReq('Summarize this'), { context }));
    const body = JSON.parse(
      (f.mock.calls[0] as unknown as [string, RequestInit])[1].body as string,
    );
    expect(body.query.parts).toBeDefined();
    const parts = body.query.parts as Array<{ text?: string }>;
    // Last part is the user's question; an earlier part carries the attached data.
    expect(parts[parts.length - 1]?.text).toBe('Summarize this');
    expect(parts.length).toBeGreaterThan(1);
  });
});
