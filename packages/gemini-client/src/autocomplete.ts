import { completeQueryUrl, GeminiClientConfig } from './config.js';
import { DeCompleteQueryResponseSchema } from './de-search-types.js';
import { defaultFetch, postJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

export interface CompleteOptions {
  /** Maximum number of suggestions to return for the QUERY suggestion type. */
  maxSuggestions?: number;
  /** Return tail suggestions when nothing matches the full query. */
  includeTailSuggestions?: boolean;
  /** Override the autocomplete query model. */
  queryModel?: string;
  signal?: AbortSignal;
}

/**
 * Discovery Engine `completionConfig:completeQuery` over the engine, called
 * client-direct as the signed-in user. Powers type-ahead in the panel composer
 * and Excel formula prompts.
 */
export class AutocompleteClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  async complete(query: string, opts: CompleteOptions = {}): Promise<string[]> {
    const url = completeQueryUrl(this.config);
    const body = buildCompleteQueryRequest(query, opts);
    const json = await postJson(url, body, this.tokens, this.fetchImpl, opts.signal);
    return mapCompleteQueryResponse(json);
  }
}

export function buildCompleteQueryRequest(
  query: string,
  opts: CompleteOptions = {},
): Record<string, unknown> {
  const out: Record<string, unknown> = { query };
  if (opts.includeTailSuggestions) out.includeTailSuggestions = true;
  if (opts.queryModel) out.queryModel = opts.queryModel;
  if (opts.maxSuggestions !== undefined) {
    out.suggestionTypeSpecs = [{ suggestionType: 'QUERY', maxSuggestions: opts.maxSuggestions }];
  }
  return out;
}

export function mapCompleteQueryResponse(json: unknown): string[] {
  const parsed = DeCompleteQueryResponseSchema.safeParse(json);
  if (!parsed.success) return [];
  const out: string[] = [];
  for (const s of parsed.data.querySuggestions ?? []) {
    if (s.suggestion && !out.includes(s.suggestion)) out.push(s.suggestion);
  }
  return out;
}
