import { z } from 'zod';

/**
 * Lenient Zod views of the Discovery Engine `v1alpha` retrieval wire types we
 * consume: `:search`, `completionConfig:completeQuery`, `groundingConfigs:check`,
 * and `rankingConfigs:rank`. As with de-types.ts these are deliberately partial +
 * passthrough — the engine emits many more fields than we depend on, and an
 * unexpected extra field must never break parsing.
 */

/* ---------------------------------------------------------------- :search --- */

export const DeSearchDocumentSchema = z
  .object({
    name: z.string().optional(),
    id: z.string().optional(),
    // title / uri usually surface here (derived from the source) or in structData.
    derivedStructData: z.record(z.unknown()).optional(),
    structData: z.record(z.unknown()).optional(),
  })
  .passthrough();

export const DeSearchResultSchema = z
  .object({
    id: z.string().optional(),
    document: DeSearchDocumentSchema.optional(),
  })
  .passthrough();

export const DeFacetValueSchema = z
  .object({
    value: z.string().optional(),
    count: z.union([z.string(), z.number()]).optional(),
  })
  .passthrough();

export const DeFacetSchema = z
  .object({
    key: z.string().optional(),
    values: z.array(DeFacetValueSchema).optional(),
    dynamicFacet: z.boolean().optional(),
  })
  .passthrough();

export const DeSearchSummarySchema = z
  .object({
    summaryText: z.string().optional(),
    summaryWithMetadata: z.object({ summary: z.string().optional() }).passthrough().optional(),
  })
  .passthrough();

export const DeSearchResponseSchema = z
  .object({
    results: z.array(DeSearchResultSchema).optional(),
    facets: z.array(DeFacetSchema).optional(),
    summary: DeSearchSummarySchema.optional(),
    totalSize: z.union([z.number(), z.string()]).optional(),
    nextPageToken: z.string().optional(),
    correctedQuery: z.string().optional(),
  })
  .passthrough();

export type DeSearchResponse = z.infer<typeof DeSearchResponseSchema>;

/* ------------------------------------------------------- :completeQuery --- */

export const DeQuerySuggestionSchema = z
  .object({
    suggestion: z.string().optional(),
    score: z.number().optional(),
  })
  .passthrough();

export const DeCompleteQueryResponseSchema = z
  .object({
    querySuggestions: z.array(DeQuerySuggestionSchema).optional(),
  })
  .passthrough();

export type DeCompleteQueryResponse = z.infer<typeof DeCompleteQueryResponseSchema>;

/* -------------------------------------------------------- :checkGrounding --- */

export const DeFactChunkSchema = z
  .object({
    source: z.string().optional(),
    uri: z.string().optional(),
    title: z.string().optional(),
    domain: z.string().optional(),
    chunkText: z.string().optional(),
    index: z.number().optional(),
  })
  .passthrough();

export const DeGroundingClaimSchema = z
  .object({
    claimText: z.string().optional(),
    score: z.number().optional(),
    startPos: z.number().optional(),
    endPos: z.number().optional(),
    citationIndices: z.array(z.number()).optional(),
    groundingCheckRequired: z.boolean().optional(),
  })
  .passthrough();

export const DeCheckGroundingResponseSchema = z
  .object({
    supportScore: z.number().optional(),
    citedChunks: z.array(DeFactChunkSchema).optional(),
    claims: z.array(DeGroundingClaimSchema).optional(),
  })
  .passthrough();

export type DeCheckGroundingResponse = z.infer<typeof DeCheckGroundingResponseSchema>;

/* --------------------------------------------------------------- :rank --- */

export const DeRankingRecordSchema = z
  .object({
    id: z.string().optional(),
    title: z.string().optional(),
    content: z.string().optional(),
    score: z.number().optional(),
  })
  .passthrough();

export const DeRankResponseSchema = z
  .object({
    records: z.array(DeRankingRecordSchema).optional(),
  })
  .passthrough();

export type DeRankResponse = z.infer<typeof DeRankResponseSchema>;
