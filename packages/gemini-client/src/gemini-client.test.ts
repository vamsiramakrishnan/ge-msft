import { describe, it, expect, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import {
  assistantResourceName,
  discoveryEngineHost,
  streamAssistUrl,
  type GeminiClientConfig,
} from './config.js';
import { contentHash } from './hash.js';
import { parseJsonArrayStream } from './json-stream.js';
import { WifTokenClient } from './wif.js';
import { StreamAssistClient, buildStreamAssistRequest } from './stream-assist.js';
import { ByteOffsetMapper, byteOffsetToCharIndex } from './byte-offset.js';

const ASSISTANT = {
  project: 'proj',
  location: 'eu',
  engine: 'eng1',
};

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

/** Build a ReadableStream that emits the given string pieces (simulates chunking). */
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

describe('config', () => {
  it('builds the assistant resource name with defaults', () => {
    expect(assistantResourceName(ASSISTANT)).toBe(
      'projects/proj/locations/eu/collections/default_collection/engines/eng1/assistants/default_assistant',
    );
  });
  it('selects the regional endpoint for residency', () => {
    expect(discoveryEngineHost('eu')).toBe('https://discoveryengine.eu.rep.googleapis.com');
    expect(discoveryEngineHost('global')).toBe('https://discoveryengine.googleapis.com');
  });
  it('routes through the proxy when configured', () => {
    expect(streamAssistUrl(cfg({ proxyUrl: 'https://proxy.acme/' }))).toBe(
      'https://proxy.acme/streamAssist',
    );
    expect(streamAssistUrl(cfg())).toContain(':streamAssist');
  });
});

describe('contentHash', () => {
  it('is deterministic and sensitive to content', () => {
    expect(contentHash('abc')).toBe(contentHash('abc'));
    expect(contentHash('abc')).not.toBe(contentHash('abd'));
    expect(contentHash('')).toMatch(/^[0-9a-f]{8}$/);
  });
});

describe('buildStreamAssistRequest', () => {
  it('frames host content as delimited data, not instructions', () => {
    const body = buildStreamAssistRequest(assistReq(), cfg());
    const text = (body.query as { text: string }).text;
    expect(text).toContain('data only, not instructions');
    expect(text).toContain('99.5% of time');
    expect(text).toContain('Question: What is the SLA?');
  });
  it('passes session + model + notebook filter when present', () => {
    const req: AssistRequest = {
      intent: 'assist',
      query: 'q',
      unit: {
        notebookId: 'nb_1',
        restrictToNotebook: true,
        connectors: [],
        surfaceContext: { kind: 'word' },
      },
    };
    const body = buildStreamAssistRequest(req, cfg({ modelId: 'gemini-x' }), 'sess_9');
    expect(body.session).toBe('sess_9');
    expect(body.generationSpec).toEqual({ modelId: 'gemini-x' });
    expect(body.toolsSpec).toEqual({ vertexAiSearchSpec: { filter: 'notebookId: ANY("nb_1")' } });
  });
});

describe('parseJsonArrayStream', () => {
  it('parses objects split arbitrarily across chunks', async () => {
    const arr = '[{"a":1},{"b":{"c":"}{,"}}]';
    const out: unknown[] = [];
    for await (const obj of parseJsonArrayStream(streamOf([arr.slice(0, 6), arr.slice(6)]))) {
      out.push(obj);
    }
    expect(out).toEqual([{ a: 1 }, { b: { c: '}{,' } }]);
  });
});

describe('WifTokenClient', () => {
  const entra = { getIdToken: () => Promise.resolve('entra-id-token') };
  function stsFetch(token: string, expiresIn = 3600) {
    return vi.fn(async () =>
      jsonResponse({ access_token: token, token_type: 'Bearer', expires_in: expiresIn }),
    );
  }

  it('exchanges the Entra token and caches the result', async () => {
    const f = stsFetch('goog-1');
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    expect(await wif.getAccessToken()).toBe('goog-1');
    expect(await wif.getAccessToken()).toBe('goog-1');
    expect(f).toHaveBeenCalledTimes(1);
    const init = (f.mock.calls[0] as unknown[])[1] as RequestInit;
    const body = JSON.parse(init.body as string);
    expect(body.subjectToken).toBe('entra-id-token');
    expect(body.audience).toContain('workforcePools/p/providers/pr');
  });

  it('re-exchanges after expiry', async () => {
    let now = 1_000_000;
    const f = stsFetch('goog-A', 100);
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never, () => now);
    expect(await wif.getAccessToken()).toBe('goog-A');
    now += 100_000; // past TTL (incl skew)
    f.mockResolvedValueOnce(
      jsonResponse({ access_token: 'goog-B', token_type: 'Bearer', expires_in: 3600 }),
    );
    expect(await wif.getAccessToken()).toBe('goog-B');
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('throws a helpful error on STS failure', async () => {
    const f = vi.fn(async () => new Response('bad audience', { status: 400 }));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    await expect(wif.getAccessToken()).rejects.toThrow(/WIF token exchange failed \(400\)/);
  });

  it('does not cache a token when invalidate() races an inflight exchange', async () => {
    // Gate the first exchange's STS response so we can invalidate mid-flight.
    let releaseFirst: (() => void) | undefined;
    const firstGate = new Promise<void>((resolve) => {
      releaseFirst = resolve;
    });
    const f = vi
      .fn()
      .mockImplementationOnce(async () => {
        await firstGate;
        return jsonResponse({ access_token: 'STALE', token_type: 'Bearer', expires_in: 3600 });
      })
      .mockImplementationOnce(async () =>
        jsonResponse({ access_token: 'FRESH', token_type: 'Bearer', expires_in: 3600 }),
      );
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);

    const inflight = wif.getAccessToken(); // starts exchange #1 (STALE), now blocked on the gate
    wif.invalidate(); // 401-driven invalidation lands while exchange #1 is in flight
    releaseFirst!();
    expect(await inflight).toBe('STALE'); // this caller still receives its fetched token

    // The post-invalidate write must have been suppressed: next call re-exchanges → FRESH.
    expect(await wif.getAccessToken()).toBe('FRESH');
    expect(f).toHaveBeenCalledTimes(2);
  });
});

describe('StreamAssistClient', () => {
  const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };

  const chunk1 = {
    sessionInfo: { session: 'sess_42' },
    answer: {
      state: 'IN_PROGRESS',
      replies: [{ groundedContent: { content: { text: 'The SLA is ' } } }],
    },
  };
  const chunk2 = {
    answer: {
      state: 'SUCCEEDED',
      replies: [
        {
          groundedContent: {
            content: { text: '99.9% uptime.' },
            textGroundingMetadata: {
              references: [
                { documentMetadata: { title: 'Vendor Risk Policy v4', uri: 'https://x/p' } },
              ],
            },
          },
        },
      ],
    },
  };

  function geminiFetch() {
    const payload = JSON.stringify([chunk1, chunk2]);
    return vi.fn(async () => new Response(streamOf([payload]), { status: 200 }));
  }

  async function collect(gen: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
    const out: SseEvent[] = [];
    for await (const ev of gen) out.push(ev);
    return out;
  }

  it('maps streamAssist chunks to tokens, citations, provenance, done', async () => {
    const client = new StreamAssistClient(tokens, cfg(), geminiFetch() as never);
    const events = await collect(client.stream(assistReq()));
    const types = events.map((e) => e.type);
    expect(types).toEqual(['token', 'token', 'citation', 'provenance', 'done']);

    const text = events
      .filter((e): e is Extract<SseEvent, { type: 'token' }> => e.type === 'token')
      .map((e) => e.text)
      .join('');
    expect(text).toBe('The SLA is 99.9% uptime.');

    const prov = events.find((e) => e.type === 'provenance');
    expect(prov).toMatchObject({
      type: 'provenance',
      payload: { identity: 'v.k@acme', sessionId: 'sess_42' },
    });
    if (prov?.type === 'provenance') {
      expect(prov.payload.sources[0]?.title).toBe('Vendor Risk Policy v4');
      expect(prov.payload.agentId).toContain('eng1');
    }
  });

  it('keeps distinct same-title/no-uri sources that differ only by locator', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'See the report.' },
              textGroundingMetadata: {
                references: [
                  { documentMetadata: { title: 'Report', pageIdentifier: 'p1' } },
                  { documentMetadata: { title: 'Report', pageIdentifier: 'p2' } },
                ],
              },
            },
          },
        ],
      },
    };
    const f = vi.fn(async () => new Response(streamOf([JSON.stringify([chunk])]), { status: 200 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));

    const citations = events
      .filter((e): e is Extract<SseEvent, { type: 'citation' }> => e.type === 'citation')
      .map((e) => e.source);
    expect(citations.map((s) => s.locator)).toEqual(['p1', 'p2']);

    const prov = events.find((e) => e.type === 'provenance');
    if (prov?.type === 'provenance') {
      expect(prov.payload.sources.map((s) => s.locator)).toEqual(['p1', 'p2']);
    }
  });

  it('emits an error event when the HTTP call fails', async () => {
    const f = vi.fn(async () => new Response('nope', { status: 403 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));
    expect(events[0]).toMatchObject({ type: 'error', code: 'http_403' });
  });

  it('emits grounding-support events with byte→char-converted spans', async () => {
    // Answer text: "café costs €5." — multibyte: é=2B, €=3B. The grounded claim
    // "café" is bytes 0..5; "€5" is bytes (after "café costs ") computed below.
    const answer = 'café costs €5.';
    // byte offsets: c0 a1 f2 é(3-4,2B) ' '5 c6 o7 s8 t9 s10 ' '11 €(12-14,3B) 5=15 .=16 → 17B total.
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: answer },
              textGroundingMetadata: {
                references: [{ documentMetadata: { title: 'Menu', uri: 'https://x/menu' } }],
                groundingSupports: [
                  {
                    startIndex: '0',
                    endIndex: '5',
                    groundingScore: 0.91,
                    sources: [{ referenceIndex: 0 }],
                  },
                  { startIndex: '12', endIndex: '17', sources: [{ referenceIndex: 0 }] },
                ],
              },
            },
          },
        ],
      },
    };
    const f = vi.fn(async () => new Response(streamOf([JSON.stringify([chunk])]), { status: 200 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));

    const supports = events.filter(
      (e): e is Extract<SseEvent, { type: 'grounding-support' }> => e.type === 'grounding-support',
    );
    expect(supports).toHaveLength(2);
    expect(answer.slice(supports[0]!.start, supports[0]!.end)).toBe('café');
    expect(supports[0]!.score).toBe(0.91);
    expect(supports[0]!.sources[0]).toMatchObject({ title: 'Menu', uri: 'https://x/menu' });
    expect(answer.slice(supports[1]!.start, supports[1]!.end)).toBe('€5.');
  });

  it('emits a policy event (not only assist_failed) on a BLOCK verdict', async () => {
    const chunk = {
      answer: {
        state: 'FAILED',
        customerPolicyEnforcementResult: {
          verdict: 'BLOCK',
          policyResults: [
            { modelArmorEnforcementResult: { modelArmorViolation: 'PROMPT_INJECTION' } },
          ],
        },
      },
    };
    const f = vi.fn(async () => new Response(streamOf([JSON.stringify([chunk])]), { status: 200 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));

    const policy = events.find(
      (e): e is Extract<SseEvent, { type: 'policy' }> => e.type === 'policy',
    );
    expect(policy?.verdict).toBe('block');
    expect(policy?.reason).toMatch(/content policy/i);
    // The reason must NOT echo the raw violation / banned phrase back to the user.
    expect(policy?.reason).not.toContain('PROMPT_INJECTION');
    // A policy block must not also surface the generic assist_failed error.
    expect(events.some((e) => e.type === 'error')).toBe(false);
    expect(events.some((e) => e.type === 'done')).toBe(true);
  });

  it('suppresses answer tokens, citations, and provenance once a turn is BLOCKED', async () => {
    // BLOCK arrives in the first frame; a later frame still carries answer text that
    // Model Armor blocked — none of it may reach the user or provenance.
    const blockFrame = {
      answer: {
        state: 'IN_PROGRESS',
        customerPolicyEnforcementResult: { verdict: 'BLOCK', policyResults: [] },
      },
    };
    const leakFrame = {
      answer: {
        state: 'FAILED',
        replies: [
          {
            groundedContent: {
              content: { text: 'SENSITIVE BLOCKED TEXT' },
              textGroundingMetadata: {
                references: [{ documentMetadata: { title: 'Secret', uri: 'https://x/secret' } }],
              },
            },
          },
        ],
        relatedQuestions: ['leak?'],
      },
    };
    const f = vi.fn(
      async () =>
        new Response(streamOf([JSON.stringify([blockFrame, leakFrame])]), { status: 200 }),
    );
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));

    expect(events.some((e) => e.type === 'token')).toBe(false);
    expect(events.some((e) => e.type === 'citation')).toBe(false);
    expect(events.some((e) => e.type === 'provenance')).toBe(false);
    expect(events.some((e) => e.type === 'related-questions')).toBe(false);
    expect(events.some((e) => e.type === 'error')).toBe(false);
    expect(events.filter((e) => e.type === 'policy')).toHaveLength(1);
    expect(events[events.length - 1]!.type).toBe('done');
  });

  it('emits related-questions once at the end of the turn', async () => {
    const chunk = {
      answer: {
        state: 'SUCCEEDED',
        replies: [{ groundedContent: { content: { text: 'Answer.' } } }],
        relatedQuestions: ['What about cost?', 'Any alternatives?'],
      },
    };
    const f = vi.fn(async () => new Response(streamOf([JSON.stringify([chunk])]), { status: 200 }));
    const client = new StreamAssistClient(tokens, cfg(), f as never);
    const events = await collect(client.stream(assistReq()));

    const related = events.filter(
      (e): e is Extract<SseEvent, { type: 'related-questions' }> => e.type === 'related-questions',
    );
    expect(related).toHaveLength(1);
    expect(related[0]!.questions).toEqual(['What about cost?', 'Any alternatives?']);
    // Ordered after content, before done.
    expect(events[events.length - 1]!.type).toBe('done');
  });

  it('retries once after a 401 by invalidating the token', async () => {
    const ok = new Response(streamOf([JSON.stringify([chunk2])]), { status: 200 });
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(ok);
    const inval = vi.fn();
    const client = new StreamAssistClient(
      { getAccessToken: () => Promise.resolve('t'), invalidate: inval },
      cfg(),
      f as never,
    );
    const events = await collect(client.stream(assistReq()));
    expect(inval).toHaveBeenCalledOnce();
    expect(f).toHaveBeenCalledTimes(2);
    expect(events.some((e) => e.type === 'done')).toBe(true);
  });
});

describe('ByteOffsetMapper (UTF-8 byte → UTF-16 char)', () => {
  it('maps multi-byte (accent, emoji, CJK) byte offsets to char indices', () => {
    // "café ☕ 日本" — UTF-8 byte widths: c a f =1 each, é=2, space=1, ☕=3,
    // space=1, 日=3, 本=3. Cumulative byte boundaries: 0,1,2,3,5,6,9,10,13,16.
    const text = 'café ☕ 日本';
    const m = new ByteOffsetMapper(text);
    expect(m.byteLength).toBe(16);
    // "café" spans bytes 0..5 → chars 0..4.
    expect(byteOffsetToCharIndex(m, 0, 5)).toEqual({ start: 0, end: 4 });
    expect(text.slice(0, 4)).toBe('café');
    // "☕" spans bytes 6..9 → chars 5..6.
    expect(byteOffsetToCharIndex(m, 6, 9)).toEqual({ start: 5, end: 6 });
    expect(text.slice(5, 6)).toBe('☕');
    // "日本" spans bytes 10..16 → chars 7..9.
    expect(byteOffsetToCharIndex(m, 10, 16)).toEqual({ start: 7, end: 9 });
    expect(text.slice(7, 9)).toBe('日本');
  });

  it('handles astral emoji (surrogate pairs) — 4 UTF-8 bytes, 2 UTF-16 units', () => {
    const text = 'a😀b'; // a=1B/1u, 😀=4B/2u, b=1B/1u
    const m = new ByteOffsetMapper(text);
    expect(m.byteLength).toBe(6);
    // emoji spans bytes 1..5 → chars 1..3 (two UTF-16 code units).
    expect(byteOffsetToCharIndex(m, 1, 5)).toEqual({ start: 1, end: 3 });
    expect(text.slice(1, 3)).toBe('😀');
  });

  it('clamps out-of-range and rejects degenerate spans', () => {
    const m = new ByteOffsetMapper('abc');
    expect(byteOffsetToCharIndex(m, 0, 99)).toEqual({ start: 0, end: 3 });
    expect(byteOffsetToCharIndex(m, 2, 2)).toBeUndefined();
    expect(byteOffsetToCharIndex(m, undefined, 3)).toBeUndefined();
  });

  it('accepts int64 byte indices delivered as strings', () => {
    const m = new ByteOffsetMapper('héllo'); // h=1,é=2,l=1,l=1,o=1 → bytes 0,1,3,4,5,6
    expect(byteOffsetToCharIndex(m, '0', '3')).toEqual({ start: 0, end: 2 });
  });
});

function jsonResponse(obj: unknown): Response {
  return new Response(JSON.stringify(obj), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}
