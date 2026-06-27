import { describe, it, expect, vi } from 'vitest';
import { type GeminiClientConfig } from './config.js';
import {
  GroundingClient,
  buildCheckGroundingRequest,
  mapCheckGroundingResponse,
} from './grounding.js';

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

describe('buildCheckGroundingRequest', () => {
  it('maps facts to factText/attributes and options to groundingSpec', () => {
    const body = buildCheckGroundingRequest(
      'The SLA is 99.9%.',
      [{ text: 'SLA is 99.9% uptime', attributes: { source: 'policy-v4' } }],
      { claimLevelScore: true, citationThreshold: 0.6 },
    );
    expect(body.answerCandidate).toBe('The SLA is 99.9%.');
    expect(body.facts).toEqual([
      { factText: 'SLA is 99.9% uptime', attributes: { source: 'policy-v4' } },
    ]);
    expect(body.groundingSpec).toEqual({ enableClaimLevelScore: true, citationThreshold: 0.6 });
  });
});

describe('mapCheckGroundingResponse', () => {
  it('produces an ergonomic shape for a threshold gate', () => {
    const out = mapCheckGroundingResponse({
      supportScore: 0.92,
      citedChunks: [{ source: '0', uri: 'https://x/p', title: 'Policy', chunkText: '99.9%' }],
      claims: [{ claimText: 'The SLA is 99.9%.', score: 0.95, citationIndices: [0] }],
    });
    expect(out.supportScore).toBe(0.92);
    expect(out.citedChunks[0]).toMatchObject({ uri: 'https://x/p', title: 'Policy' });
    expect(out.claims?.[0]).toMatchObject({ score: 0.95, citationIndices: [0] });
  });
  it('defaults supportScore to 0 on a malformed response', () => {
    expect(mapCheckGroundingResponse({ junk: 1 })).toEqual({ supportScore: 0, citedChunks: [] });
  });
});

describe('GroundingClient', () => {
  it('POSTs to groundingConfigs:check and maps the response', async () => {
    const f = jsonFetch({ supportScore: 0.8, citedChunks: [] });
    const client = new GroundingClient(tokens, cfg(), f as never);
    const out = await client.check('cand', [{ text: 'fact' }], { citationThreshold: 0.5 });
    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toContain('groundingConfigs/default_grounding_config:check');
    expect(url).not.toContain('/engines/');
    expect(JSON.parse(init.body as string).answerCandidate).toBe('cand');
    expect(out.supportScore).toBe(0.8);
  });
});
