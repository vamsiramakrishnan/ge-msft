import { z } from 'zod';
import { SourceVersionSchema } from './analysis.js';
import { ChangeIdSchema } from './brand.js';

const hash = z.string().regex(/^sha256:[a-f0-9]{64}$/);
const sources = z.array(SourceVersionSchema).max(16);
const identity = z.string().min(1).max(256);

/** Captured from the admitted effect, not a model-authored locator or persisted approval. */
export const DurableInspectionEvidenceSchema = z.discriminatedUnion('kind', [
  z
    .object({
      version: z.literal(1),
      kind: z.literal('cells'),
      operationKind: z.literal('write-cells'),
      changeId: ChangeIdSchema,
      documentId: identity,
      objectId: identity,
      locator: z.string().min(1).max(1024),
      sources,
    })
    .strict(),
  z
    .object({
      version: z.literal(1),
      kind: z.literal('text'),
      operationKind: z.enum(['fill-content-control', 'set-shape-text']),
      changeId: ChangeIdSchema,
      surface: z.enum(['word', 'powerpoint']),
      documentId: hash,
      objectId: identity,
      parentObjectId: identity.optional(),
      sources,
    })
    .strict(),
]);
export type DurableInspectionEvidence = z.infer<typeof DurableInspectionEvidenceSchema>;
