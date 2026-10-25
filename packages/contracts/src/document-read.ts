import { z } from 'zod';

/** A host read page, not a model prompt. EOF is explicit; an outline is never a page. */
export const DocumentPageSchema = z.object({
  scope: z.string().min(1).max(500),
  units: z
    .array(
      z.object({
        id: z.string().min(1).max(1000),
        text: z.string(),
        parentId: z.string().min(1).max(1000).optional(),
        relatedIds: z.array(z.string().min(1).max(1000)).max(100).optional(),
        /** Adapter-declared relation coverage only; absence is unknown, never write authority. */
        relationsComplete: z.boolean().optional(),
      }),
    )
    .max(1000),
  nextCursor: z.string().min(1).max(4000).optional(),
  limitations: z.array(z.string().max(1000)).max(20),
});
export type DocumentPage = z.infer<typeof DocumentPageSchema>;
export interface DocumentReadOptions {
  cursor?: string;
  signal?: AbortSignal;
}

/** Explicit SDK scope avoids depending on natural-language inference. */
export type DocumentScope = 'attached' | 'document';

/** Convenience routing only. Extraction and coverage are independent of prompt wording. */
export function requestsDocumentSummary(query: string): boolean {
  const summary = /\b(summari[sz]e|summary|summaries|recap|overview|synthesi[sz]e)\b/i.test(query);
  const whole =
    /\b(entire|whole|complete|full)\s+(?:current\s+)?(doc(?:ument)?|deck|presentation|workbook|spreadsheet)\b|\ball\s+(?:\d+\s+)?(slides|sheets|tabs|worksheets|pages)\b/i.test(
      query,
    );
  const document =
    /\b(?:summari[sz]e|summary\s+of|recap\s+of|overview\s+of)\s+(?:(?:the|this|current)\s+)?(?:document|doc|deck|presentation|workbook|spreadsheet)\b/i.test(
      query,
    );
  return summary && (whole || document);
}

/** Reject invented/negative cursors before invoking a host API. */
export function documentOffset(cursor: string | undefined): number {
  if (cursor === undefined) return 0;
  if (!/^\d+$/.test(cursor)) throw new Error('Invalid document cursor. Start a new read.');
  const offset = Number(cursor);
  if (!Number.isSafeInteger(offset) || offset > 10_000_000)
    throw new Error('Invalid document cursor. Start a new read.');
  return offset;
}
