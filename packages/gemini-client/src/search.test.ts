import { describe, it, expect, vi } from 'vitest';
import {
  searchUrl,
  completeQueryUrl,
  checkGroundingUrl,
  rankUrl,
  type GeminiClientConfig,
} from './config.js';
import {
  SearchClient,
  buildSearchRequest,
  searchHitToContextRef,
  searchHitToResolvedContext,
  type SearchHit,
} from './search.js';

const ASSISTANT = { project: 'proj', location: 'eu', engine: 'eng1' };

function cfg(overrides: Partial<GeminiClientConfig> = {}): GeminiClientConfig {
  return { assistant: ASSISTANT, identity: 'v.k@acme', ...overrides };
}

function jsonFetch(obj: unknown, status = 200) {
  return vi.fn(
    async () =>
      new Response(JSON.stringify(obj), {
        status,
        headers: { 'Content-Type': 'application/json' },
      }),
  );
}

const tokens = { getAccessToken: () => Promise.resolve('goog-token'), invalidate: vi.fn() };

const DOC_NAME =
  'projects/proj/locations/eu/collections/default_collection/dataStores/ds1/branches/0/documents/d1';

const searchResponse = {
  results: [
    {
      id: 'd1',
      document: {
        name: DOC_NAME,
        id: 'd1',
        derivedStructData: {
          title: 'Vendor Risk Policy v4',
          link: 'https://x/p',
          snippets: [{ snippet: 'SLA is 99.9%', snippet_status: 'SUCCESS' }],
        },
        structData: { owner: 'legal' },
      },
    },
  ],
  facets: [
    {
      key: 'category',
      values: [
        { value: 'policy', count: '12' },
        { value: 'contract', count: 3 },
      ],
    },
  ],
  summary: { summaryText: 'Policies summarized.' },
  totalSize: 42,
  nextPageToken: 'page2',
  correctedQuery: 'sla policy',
};

describe('search config URLs', () => {
  it('builds the engine-scoped serving config search URL with default_search', () => {
    expect(searchUrl(cfg())).toBe(
      'https://discoveryengine.eu.rep.googleapis.com/v1alpha/projects/proj/locations/eu/collections/default_collection/engines/eng1/servingConfigs/default_search:search',
    );
  });
  it('builds the completeQuery URL under completionConfig', () => {
    expect(completeQueryUrl(cfg())).toContain('/engines/eng1/completionConfig:completeQuery');
  });
  it('scopes checkGrounding to project+location (not the engine)', () => {
    expect(checkGroundingUrl(cfg())).toBe(
      'https://discoveryengine.eu.rep.googleapis.com/v1alpha/projects/proj/locations/eu/groundingConfigs/default_grounding_config:check',
    );
    expect(checkGroundingUrl(cfg())).not.toContain('/engines/');
  });
  it('scopes rank to project+location (not the engine)', () => {
    expect(rankUrl(cfg())).toBe(
      'https://discoveryengine.eu.rep.googleapis.com/v1alpha/projects/proj/locations/eu/rankingConfigs/default_ranking_config:rank',
    );
    expect(rankUrl(cfg())).not.toContain('/engines/');
  });
  it('routes through the proxy when configured', () => {
    expect(searchUrl(cfg({ proxyUrl: 'https://proxy.acme/' }))).toBe('https://proxy.acme/search');
    expect(rankUrl(cfg({ proxyUrl: 'https://proxy.acme' }))).toBe('https://proxy.acme/rank');
  });
});

describe('buildSearchRequest', () => {
  it('maps options onto the DE SearchRequest body shape', () => {
    const body = buildSearchRequest({
      query: 'sla',
      pageSize: 5,
      pageToken: 'tok',
      filter: 'category: ANY("policy")',
      orderBy: 'indexTime desc',
      facetSpecs: ['category', 'owner'],
      boostSpec: { conditionBoostSpecs: [{ condition: 'x', boost: 1 }] },
      dataStoreSpecs: [{ dataStore: 'projects/p/.../dataStores/ds1', filter: 'a: ANY("b")' }],
      summary: true,
      snippets: true,
    });
    expect(body.query).toBe('sla');
    expect(body.pageSize).toBe(5);
    expect(body.pageToken).toBe('tok');
    expect(body.filter).toBe('category: ANY("policy")');
    expect(body.orderBy).toBe('indexTime desc');
    expect(body.facetSpecs).toEqual([
      { facetKey: { key: 'category' } },
      { facetKey: { key: 'owner' } },
    ]);
    expect(body.dataStoreSpecs).toEqual([
      { dataStore: 'projects/p/.../dataStores/ds1', filter: 'a: ANY("b")' },
    ]);
    expect(body.contentSearchSpec).toEqual({
      snippetSpec: { returnSnippet: true },
      summarySpec: { includeCitations: true },
    });
  });
  it('omits absent specs', () => {
    const body = buildSearchRequest({ query: 'q' });
    expect(body).toEqual({ query: 'q' });
  });
});

describe('SearchClient', () => {
  it('POSTs the request body and maps the response into a clean result', async () => {
    const f = jsonFetch(searchResponse);
    const client = new SearchClient(tokens, cfg(), f as never);
    const out = await client.search({ query: 'sla', facetSpecs: ['category'], snippets: true });

    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toContain('servingConfigs/default_search:search');
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer goog-token');
    const sent = JSON.parse(init.body as string);
    expect(sent.query).toBe('sla');
    expect(sent.facetSpecs).toEqual([{ facetKey: { key: 'category' } }]);

    expect(out.results).toHaveLength(1);
    expect(out.results[0]).toMatchObject({
      id: 'd1',
      documentName: DOC_NAME,
      title: 'Vendor Risk Policy v4',
      uri: 'https://x/p',
      snippet: 'SLA is 99.9%',
      structData: { owner: 'legal' },
    });
    expect(out.facets).toEqual([
      {
        key: 'category',
        values: [
          { value: 'policy', count: 12 },
          { value: 'contract', count: 3 },
        ],
      },
    ]);
    expect(out.summary).toBe('Policies summarized.');
    expect(out.totalSize).toBe(42);
    expect(out.nextPageToken).toBe('page2');
    expect(out.correctedQuery).toBe('sla policy');
  });

  it('retries once after a 401 by invalidating the token', async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(
        new Response(JSON.stringify(searchResponse), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }),
      );
    const inval = vi.fn();
    const client = new SearchClient(
      { getAccessToken: () => Promise.resolve('t'), invalidate: inval },
      cfg(),
      f as never,
    );
    const out = await client.search({ query: 'sla' });
    expect(inval).toHaveBeenCalledOnce();
    expect(f).toHaveBeenCalledTimes(2);
    expect(out.results).toHaveLength(1);
  });

  it('throws on a non-401 HTTP error', async () => {
    const f = jsonFetch({}, 403);
    const client = new SearchClient(tokens, cfg(), f as never);
    await expect(client.search({ query: 'x' })).rejects.toThrow(/403/);
  });
});

describe('search → context bridge (reference over inline)', () => {
  const hit: SearchHit = {
    id: 'd1',
    documentName: DOC_NAME,
    title: 'Vendor Risk Policy v4',
    uri: 'https://x/p',
    snippet: 'SLA is 99.9%',
  };

  it('searchHitToContextRef yields an indexed-document ref with a preview', () => {
    const ref = searchHitToContextRef(hit);
    expect(ref.kind).toBe('indexed-document');
    expect(ref.id).toBe(`indexed:${DOC_NAME}`);
    expect(ref.title).toBe('Vendor Risk Policy v4');
    expect(ref.preview).toBe('SLA is 99.9%');
  });

  it('searchHitToResolvedContext references the doc by name, not inlined text', () => {
    const resolved = searchHitToResolvedContext(hit);
    expect(resolved.value).toEqual({
      as: 'indexed-document',
      documentName: DOC_NAME,
      title: 'Vendor Risk Policy v4',
      uri: 'https://x/p',
    });
    expect(resolved.ref.kind).toBe('indexed-document');
  });
});
