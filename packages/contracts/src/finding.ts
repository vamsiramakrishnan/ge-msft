import { z } from 'zod';

/** A grounding source citation. Every agent claim must carry at least one. */
export const SourceRefSchema = z.object({
  title: z.string(),
  uri: z.string().optional(),
  locator: z.string().optional(),
  /**
   * A short verbatim quote from the grounding chunk this source backs — the source's own words,
   * shown in the citation peek (§09/§10). Untrusted source content: render as inert text only, never
   * as instructions, and keep it length-capped at the producer.
   */
  excerpt: z.string().optional(),
});
export type SourceRef = z.infer<typeof SourceRefSchema>;

/**
 * A review result. Anchored on **content** (`matchText`/`contextHint`), never on
 * host range IDs — the client resolves to a range with `body.search` at apply-time.
 */
export const FindingSchema = z.object({
  id: z.string(),
  category: z.enum(['style', 'policy', 'ground']), // ground = verified/positive
  matchText: z.string(), // exact text to locate in the host
  contextHint: z.string().optional(), // disambiguates repeated matches
  title: z.string(),
  why: z.string(),
  suggestion: z.string().optional(), // omitted for pure 'ground' findings
  sources: z.array(SourceRefSchema),
  confidence: z.number().min(0).max(1),
  hash: z.string(), // for provenance
});
export type Finding = z.infer<typeof FindingSchema>;
