import type { ContextRef, ResolvedContext } from '@ge/contracts';
import { GeminiClientConfig, searchUrl } from './config.js';
import { DeSearchResponseSchema } from './de-search-types.js';
import { defaultFetch, postJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/** A single matched indexed document, flattened for UI + the context bridge. */
export interface SearchHit {
  id: string;
  /** Full DE resource name: projects/.../dataStores/.../documents/... */
  documentName: string;
  title?: string;
  uri?: string;
  snippet?: string;
  structData?: Record<string, unknown>;
}

export interface FacetValue {
  value: string;
  count?: number;
}

export interface Facet {
  key: string;
  values: FacetValue[];
}

export interface SearchResult {
  results: SearchHit[];
  facets: Facet[];
  summary?: string;
  totalSize?: number;
  nextPageToken?: string;
  correctedQuery?: string;
}

/** A `dataStoreSpecs[]` entry scoping a sub-search to one connector/data store. */
export interface DataStoreSpec {
  /** Full resource name: projects/.../collections/.../dataStores/{id} */
  dataStore: string;
  filter?: string;
}

export interface SearchOptions {
  pageSize?: number;
  pageToken?: string;
  filter?: string;
  orderBy?: string;
  /** Facet keys to compute (document fields). */
  facetSpecs?: string[];
  /** Raw DE boostSpec passthrough. */
  boostSpec?: Record<string, unknown>;
  /** Scope the search to specific data stores (connectors). */
  dataStoreSpecs?: DataStoreSpec[];
  /** Include a generated summary (maps to contentSearchSpec.summarySpec). */
  summary?: boolean;
  /** Include snippets (maps to contentSearchSpec.snippetSpec.returnSnippet). */
  snippets?: boolean;
  signal?: AbortSignal;
}

export interface SearchRequest extends SearchOptions {
  query: string;
}

/**
 * Standard Discovery Engine `:search` over the engine-scoped serving config, called
 * client-direct as the signed-in user. Faceted/filtered retrieval that powers source
 * pickers, entity cards, and discovery of indexed docs to attach to a session.
 */
export class SearchClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  async search(req: SearchRequest): Promise<SearchResult> {
    const url = searchUrl(this.config);
    const body = buildSearchRequest(req);
    const json = await postJson(url, body, this.tokens, this.fetchImpl, req.signal);
    return mapSearchResponse(json);
  }
}

/** Map our ergonomic request onto a Discovery Engine SearchRequest body. */
export function buildSearchRequest(req: SearchRequest): Record<string, unknown> {
  const out: Record<string, unknown> = { query: req.query };
  if (req.pageSize !== undefined) out.pageSize = req.pageSize;
  if (req.pageToken) out.pageToken = req.pageToken;
  if (req.filter) out.filter = req.filter;
  if (req.orderBy) out.orderBy = req.orderBy;
  if (req.facetSpecs && req.facetSpecs.length > 0) {
    out.facetSpecs = req.facetSpecs.map((key) => ({ facetKey: { key } }));
  }
  if (req.boostSpec) out.boostSpec = req.boostSpec;
  if (req.dataStoreSpecs && req.dataStoreSpecs.length > 0) {
    out.dataStoreSpecs = req.dataStoreSpecs.map((s) => ({
      dataStore: s.dataStore,
      ...(s.filter ? { filter: s.filter } : {}),
    }));
  }
  const contentSearchSpec: Record<string, unknown> = {};
  if (req.snippets) contentSearchSpec.snippetSpec = { returnSnippet: true };
  if (req.summary) contentSearchSpec.summarySpec = { includeCitations: true };
  if (Object.keys(contentSearchSpec).length > 0) out.contentSearchSpec = contentSearchSpec;
  return out;
}

export function mapSearchResponse(json: unknown): SearchResult {
  const parsed = DeSearchResponseSchema.safeParse(json);
  if (!parsed.success) {
    return { results: [], facets: [] };
  }
  const data = parsed.data;

  const results: SearchHit[] = [];
  for (const r of data.results ?? []) {
    const doc = r.document;
    if (!doc) continue;
    const documentName = doc.name ?? doc.id ?? r.id;
    if (!documentName) continue;
    const derived = doc.derivedStructData ?? {};
    const struct = doc.structData ?? {};
    const title = firstString(derived.title, struct.title);
    const uri = firstString(derived.link, derived.uri, struct.uri, struct.link);
    const snippet = extractSnippet(derived.snippets) ?? firstString(derived.snippet);
    const hit: SearchHit = {
      id: doc.id ?? r.id ?? documentName,
      documentName,
      ...(title ? { title } : {}),
      ...(uri ? { uri } : {}),
      ...(snippet ? { snippet } : {}),
      ...(doc.structData ? { structData: doc.structData } : {}),
    };
    results.push(hit);
  }

  const facets: Facet[] = (data.facets ?? [])
    .filter((f): f is typeof f & { key: string } => typeof f.key === 'string')
    .map((f) => ({
      key: f.key,
      values: (f.values ?? [])
        .filter((v): v is typeof v & { value: string } => typeof v.value === 'string')
        .map((v) => ({
          value: v.value,
          ...(v.count !== undefined ? { count: Number(v.count) } : {}),
        })),
    }));

  const summary = data.summary?.summaryText ?? data.summary?.summaryWithMetadata?.summary;
  const totalSize = data.totalSize !== undefined ? Number(data.totalSize) : undefined;

  return {
    results,
    facets,
    ...(summary ? { summary } : {}),
    ...(totalSize !== undefined ? { totalSize } : {}),
    ...(data.nextPageToken ? { nextPageToken: data.nextPageToken } : {}),
    ...(data.correctedQuery ? { correctedQuery: data.correctedQuery } : {}),
  };
}

/**
 * Reference-over-inline bridge: turn a discovered indexed document into a lightweight
 * `ContextRef` the UI can list/attach. The doc is referenced by name, not inlined.
 */
export function searchHitToContextRef(hit: SearchHit): ContextRef {
  return {
    id: `indexed:${hit.documentName}`,
    kind: 'indexed-document',
    surface: 'word',
    title: hit.title ?? hit.uri ?? hit.documentName,
    ...(hit.snippet ? { preview: hit.snippet } : {}),
  };
}

/**
 * Reference-over-inline bridge: turn a discovered indexed document into a
 * `ResolvedContext` whose value attaches to a session as an indexed-document
 * *reference* (documentName), never as inlined text.
 */
export function searchHitToResolvedContext(hit: SearchHit): ResolvedContext {
  return {
    ref: searchHitToContextRef(hit),
    value: {
      as: 'indexed-document',
      documentName: hit.documentName,
      ...(hit.title ? { title: hit.title } : {}),
      ...(hit.uri ? { uri: hit.uri } : {}),
    },
  };
}

function firstString(...candidates: unknown[]): string | undefined {
  for (const c of candidates) {
    if (typeof c === 'string' && c.length > 0) return c;
  }
  return undefined;
}

/** derivedStructData.snippets is typically an array of { snippet, snippet_status }. */
function extractSnippet(snippets: unknown): string | undefined {
  if (!Array.isArray(snippets)) return undefined;
  for (const s of snippets) {
    if (s && typeof s === 'object' && 'snippet' in s) {
      const snippet = (s as { snippet: unknown }).snippet;
      if (typeof snippet === 'string' && snippet.length > 0) return snippet;
    }
  }
  return undefined;
}
