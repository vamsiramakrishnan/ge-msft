import { GeminiClientConfig, rankUrl } from './config.js';
import { DeRankResponseSchema } from './de-search-types.js';
import { defaultFetch, postJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/** A record to rerank. At least one of title/content must be set. */
export interface RankRecord {
  id: string;
  title?: string;
  content?: string;
}

export interface RankedRecord extends RankRecord {
  score: number;
}

export interface RankOptions {
  /** Number of top results to return; unset/<=0 returns all. */
  topN?: number;
  /** Ranking model id, e.g. `semantic-ranker-512@latest`. */
  model?: string;
  signal?: AbortSignal;
}

/**
 * Discovery Engine `rankingConfigs:rank` (project+location scoped), called
 * client-direct as the signed-in user. Semantic rerank of a candidate list against a
 * query (e.g. ordering review findings or entity matches).
 */
export class RankClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  async rank(
    query: string,
    records: RankRecord[],
    opts: RankOptions = {},
  ): Promise<RankedRecord[]> {
    const url = rankUrl(this.config);
    const body = buildRankRequest(query, records, opts);
    const json = await postJson(url, body, this.tokens, this.fetchImpl, opts.signal);
    return mapRankResponse(json);
  }
}

export function buildRankRequest(
  query: string,
  records: RankRecord[],
  opts: RankOptions = {},
): Record<string, unknown> {
  const out: Record<string, unknown> = {
    query,
    records: records.map((r) => ({
      id: r.id,
      ...(r.title !== undefined ? { title: r.title } : {}),
      ...(r.content !== undefined ? { content: r.content } : {}),
    })),
  };
  if (opts.topN !== undefined) out.topN = opts.topN;
  if (opts.model) out.model = opts.model;
  return out;
}

export function mapRankResponse(json: unknown): RankedRecord[] {
  const parsed = DeRankResponseSchema.safeParse(json);
  if (!parsed.success) return [];
  return (parsed.data.records ?? [])
    .filter((r): r is typeof r & { id: string } => typeof r.id === 'string')
    .map((r) => ({
      id: r.id,
      ...(r.title !== undefined ? { title: r.title } : {}),
      ...(r.content !== undefined ? { content: r.content } : {}),
      score: r.score ?? 0,
    }));
}
