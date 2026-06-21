import { z } from 'zod';
import { SourceRefSchema } from './finding.js';

/**
 * Persisted into the host's durable metadata (Word/PPT custom XML, Excel cell/
 * entity metadata, OneNote page metadata, Teams recap card). Every agent write
 * carries one so changes stay traceable and reversible.
 */
export const ProvenancePayloadSchema = z.object({
  agentId: z.string(), // e.g. "review-agent@v2"
  identity: z.string(), // signed-in user, e.g. "v.k@acme"
  timestamp: z.string(), // ISO 8601
  sources: z.array(SourceRefSchema),
  contentHash: z.string(), // hash of the generated/edited content
  sessionId: z.string().optional(), // StreamAssist session, for resume
});
export type ProvenancePayload = z.infer<typeof ProvenancePayloadSchema>;
