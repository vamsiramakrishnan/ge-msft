import { describe, it, expect, vi } from 'vitest';
import { type GeminiClientConfig } from './config.js';
import { RankClient, buildRankRequest, mapRankResponse } from './ranking.js';

const ASSISTANT = { project: 'proj', location: 'eu', engine: 'eng1' };
function cfg(overrides: Partial<GeminiClientConfig> = {}): GeminiClientConfig {
  return { assistant: ASSISTANT, identity: 'v.k@acme', ...overrides };
}

const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };

function jsonFetch(obj: unknown) {
  return vi.fn(
    async () =>
      new Response(JSON.stringify(obj), {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
      }),
  );
}

describe('buildRankRequest', () => {
  it('maps records, topN, and model', () => {
    const body = buildRankRequest(
      'best match',
      [
        { id: 'a', title: 'A', content: 'alpha' },
        { id: 'b', content: 'beta' },
      ],
      { topN: 1, model: 'semantic-ranker-512@latest' },
    );
    expect(body.query).toBe('best match');
    expect(body.topN).toBe(1);
    expect(body.model).toBe('semantic-ranker-512@latest');
    expect(body.records).toEqual([
      { id: 'a', title: 'A', content: 'alpha' },
      { id: 'b', content: 'beta' },
    ]);
  });
});

describe('mapRankResponse', () => {
  it('returns records sorted as the API ordered them, with scores', () => {
    const out = mapRankResponse({
      records: [
        { id: 'b', content: 'beta', score: 0.9 },
        { id: 'a', title: 'A', score: 0.2 },
      ],
    });
    expect(out.map((r) => r.id)).toEqual(['b', 'a']);
    expect(out[0]).toMatchObject({ id: 'b', score: 0.9 });
  });
  it('tolerates a malformed response', () => {
    expect(mapRankResponse({ x: 1 })).toEqual([]);
  });
});

describe('RankClient', () => {
  it('POSTs to rankingConfigs:rank (project+location scoped)', async () => {
    const f = jsonFetch({ records: [{ id: 'a', score: 0.5 }] });
    const client = new RankClient(tokens, cfg(), f as never);
    const out = await client.rank('q', [{ id: 'a', content: 'x' }]);
    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toContain('rankingConfigs/default_ranking_config:rank');
    expect(url).not.toContain('/engines/');
    expect(JSON.parse(init.body as string).query).toBe('q');
    expect(out).toEqual([{ id: 'a', score: 0.5 }]);
  });
});
