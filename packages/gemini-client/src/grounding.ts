import { checkGroundingUrl, GeminiClientConfig } from './config.js';
import { DeCheckGroundingResponseSchema } from './de-search-types.js';
import { postJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/** A fact to ground an answer candidate against. Carried as data, never instructions. */
export interface GroundingFact {
  text: string;
  attributes?: Record<string, string>;
}

export interface CheckGroundingOptions {
  /** Enable per-claim grounding scores in the response. */
  claimLevelScore?: boolean;
  /** Threshold in [0,1] for whether a fact must be cited for a claim. */
  citationThreshold?: number;
  signal?: AbortSignal;
}

export interface CitedChunk {
  source?: string;
  uri?: string;
  title?: string;
  chunkText?: string;
}

export interface GroundingClaim {
  claimText?: string;
  score?: number;
  startPos?: number;
  endPos?: number;
  citationIndices?: number[];
  groundingCheckRequired?: boolean;
}

export interface GroundingResult {
  /** Overall support score in [0,1]; backs a yes/no threshold gate. */
  supportScore: number;
  citedChunks: CitedChunk[];
  claims?: GroundingClaim[];
}

/**
 * Discovery Engine `groundingConfigs:check` (project+location scoped), called
 * client-direct as the signed-in user. Backs the on-send / pre-actuation grounding
 * gate: validate an agent rewrite against facts before applying a tracked change.
 */
export class GroundingClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = fetch,
  ) {}

  async check(
    answerCandidate: string,
    facts: GroundingFact[],
    opts: CheckGroundingOptions = {},
  ): Promise<GroundingResult> {
    const url = checkGroundingUrl(this.config);
    const body = buildCheckGroundingRequest(answerCandidate, facts, opts);
    const json = await postJson(url, body, this.tokens, this.fetchImpl, opts.signal);
    return mapCheckGroundingResponse(json);
  }
}

export function buildCheckGroundingRequest(
  answerCandidate: string,
  facts: GroundingFact[],
  opts: CheckGroundingOptions = {},
): Record<string, unknown> {
  const out: Record<string, unknown> = {
    answerCandidate,
    facts: facts.map((f) => ({
      factText: f.text,
      ...(f.attributes ? { attributes: f.attributes } : {}),
    })),
  };
  const groundingSpec: Record<string, unknown> = {};
  if (opts.claimLevelScore) groundingSpec.enableClaimLevelScore = true;
  if (opts.citationThreshold !== undefined)
    groundingSpec.citationThreshold = opts.citationThreshold;
  if (Object.keys(groundingSpec).length > 0) out.groundingSpec = groundingSpec;
  return out;
}

export function mapCheckGroundingResponse(json: unknown): GroundingResult {
  const parsed = DeCheckGroundingResponseSchema.safeParse(json);
  if (!parsed.success) return { supportScore: 0, citedChunks: [] };
  const data = parsed.data;

  const citedChunks: CitedChunk[] = (data.citedChunks ?? []).map((c) => ({
    ...(c.source !== undefined ? { source: c.source } : {}),
    ...(c.uri !== undefined ? { uri: c.uri } : {}),
    ...(c.title !== undefined ? { title: c.title } : {}),
    ...(c.chunkText !== undefined ? { chunkText: c.chunkText } : {}),
  }));

  const claims: GroundingClaim[] | undefined = data.claims?.map((c) => ({
    ...(c.claimText !== undefined ? { claimText: c.claimText } : {}),
    ...(c.score !== undefined ? { score: c.score } : {}),
    ...(c.startPos !== undefined ? { startPos: c.startPos } : {}),
    ...(c.endPos !== undefined ? { endPos: c.endPos } : {}),
    ...(c.citationIndices !== undefined ? { citationIndices: c.citationIndices } : {}),
    ...(c.groundingCheckRequired !== undefined
      ? { groundingCheckRequired: c.groundingCheckRequired }
      : {}),
  }));

  return {
    supportScore: data.supportScore ?? 0,
    citedChunks,
    ...(claims ? { claims } : {}),
  };
}
