import { z } from 'zod';
import { DocumentPageSchema } from './document-read.js';

export const SourceLimitsSchema = z.object({
  maxPages: z.number().int().positive().max(10000).default(10000),
  maxUnits: z.number().int().positive().max(1000000).default(1000000),
  maxBytes: z
    .number()
    .int()
    .positive()
    .max(256 * 1024 * 1024)
    .default(64 * 1024 * 1024),
});
export type SourceLimits = z.infer<typeof SourceLimitsSchema>;
export const EvidenceSnapshotSchema = z.object({
  id: z.string().regex(/^[a-f0-9]{64}$/),
  sourceId: z.string().min(1).max(1000),
  scope: z.string().min(1).max(500),
  pageHashes: z.array(z.string()).max(10000),
  unitCount: z.number().int().nonnegative().max(1000000),
  sourceBytes: z.number().int().nonnegative(),
  limitations: z.array(z.string().max(1000)).max(40),
  complete: z.literal(true),
  /** Authenticated ordinal index. Older snapshots without it use full-validation navigation. */
  pageUnitCounts: z.array(z.number().int().nonnegative().max(1000)).max(10000).optional(),
  pages: z.array(DocumentPageSchema).max(10000),
});
export type EvidenceSnapshot = z.infer<typeof EvidenceSnapshotSchema>;
export type SourceReceipt = Omit<EvidenceSnapshot, 'pages'>;
export const ContextProjectionRequestSchema = z.object({
  query: z.string().max(8192).default(''),
  requiredIds: z.array(z.string().min(1).max(1000)).max(1000).default([]),
  neighbors: z.number().int().min(0).max(10).default(1),
  relationDepth: z.number().int().min(0).max(4).default(1),
  maxBytes: z
    .number()
    .int()
    .min(512)
    .max(96 * 1024)
    .default(24 * 1024),
});
export type ContextProjectionRequest = z.input<typeof ContextProjectionRequestSchema>;
export interface EvidenceReference {
  id: string;
  hash: string;
  path: string;
}
export interface ContextProjection {
  snapshotId: string;
  sourceId: string;
  scope: string;
  sourceComplete: true;
  projectionComplete: boolean;
  selection: 'required-structure-lexical' | 'required-structure-custom';
  omittedUnits: number;
  unresolvedRelations: number;
  limitations: string[];
  units: Array<EvidenceReference & { text: string }>;
}
