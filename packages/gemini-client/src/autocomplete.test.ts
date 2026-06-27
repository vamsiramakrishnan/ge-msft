import { describe, it, expect, vi } from 'vitest';
import { type GeminiClientConfig } from './config.js';
import {
  AutocompleteClient,
  buildCompleteQueryRequest,
  mapCompleteQueryResponse,
} from './autocomplete.js';

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

describe('buildCompleteQueryRequest', () => {
  it('sends the query and maps maxSuggestions to a QUERY suggestionTypeSpec', () => {
    const body = buildCompleteQueryRequest('sl', {
      maxSuggestions: 5,
      includeTailSuggestions: true,
      queryModel: 'm1',
    });
    expect(body.query).toBe('sl');
    expect(body.includeTailSuggestions).toBe(true);
    expect(body.queryModel).toBe('m1');
    expect(body.suggestionTypeSpecs).toEqual([{ suggestionType: 'QUERY', maxSuggestions: 5 }]);
  });
});

describe('mapCompleteQueryResponse', () => {
  it('dedupes query suggestions into a string[]', () => {
    expect(
      mapCompleteQueryResponse({
        querySuggestions: [
          { suggestion: 'sla policy', score: 0.9 },
          { suggestion: 'sla report' },
          { suggestion: 'sla policy' },
        ],
      }),
    ).toEqual(['sla policy', 'sla report']);
  });
  it('tolerates a malformed response', () => {
    expect(mapCompleteQueryResponse({ unexpected: true })).toEqual([]);
  });
});

describe('AutocompleteClient', () => {
  it('POSTs to completionConfig:completeQuery and returns suggestions', async () => {
    const f = jsonFetch({ querySuggestions: [{ suggestion: 'sla policy' }] });
    const client = new AutocompleteClient(tokens, cfg(), f as never);
    const out = await client.complete('sl', { maxSuggestions: 3 });
    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toContain('completionConfig:completeQuery');
    expect(JSON.parse(init.body as string).query).toBe('sl');
    expect(out).toEqual(['sla policy']);
  });
});
