import { z } from 'zod';
import { AnchorSchema } from './anchor.js';
import { ContextKindSchema, SurfaceSchema } from './context.js';

/**
 * The ambient `<doc_state>` snapshot (ADR-0003, Layer B element 1).
 *
 * A compact, structured description of the active document, injected each turn so the
 * model knows the document's **shape** — its outline, the sheets/slides/tables present,
 * the current selection, named ranges, and comments — without having to read the whole
 * file. It is a single surface-agnostic envelope (optional fields, not a hard per-surface
 * union) produced by each bridge's `capture` and built in `@ge/content`. Like all host
 * content it is **untrusted data**: rendered wrapped (never as instructions) and screened
 * by Model Armor at the engine. See docs/ADR-0003-context-construction.md §Layer-B.
 */

/** One heading/section entry in the document outline. */
export const DocStateOutlineEntrySchema = z
  .object({
    level: z.number().int().nonnegative(), // heading depth (1..6; 0 ⇒ unleveled)
    text: z.string(), // the heading text
    anchor: AnchorSchema.optional(), // where to re-find this heading in the host
  })
  .strict();
export type DocStateOutlineEntry = z.infer<typeof DocStateOutlineEntrySchema>;

/** One present host object (sheet/slide/table/section) the model should know exists. */
export const DocStateInventoryEntrySchema = z
  .object({
    kind: ContextKindSchema, // sheet / slide / table / page / ...
    id: z.string(), // stable id (e.g. "slide:4", "range:Sheet1", "table:2")
    title: z.string(), // human label
    summary: z.string().optional(), // one-line description (e.g. table dims)
  })
  .strict();
export type DocStateInventoryEntry = z.infer<typeof DocStateInventoryEntrySchema>;

/** The current selection, if any. */
export const DocStateSelectionSchema = z
  .object({
    kind: ContextKindSchema,
    title: z.string(), // human label ("Selection", "A1:D9", "Slide 4")
    preview: z.string().optional(), // short snippet of selected content
  })
  .strict();
export type DocStateSelection = z.infer<typeof DocStateSelectionSchema>;

/** A named range (Excel). */
export const DocStateNamedRangeSchema = z
  .object({
    name: z.string(),
    range: z.string(), // A1 address (e.g. "Sheet1!$A$1:$D$9")
  })
  .strict();
export type DocStateNamedRange = z.infer<typeof DocStateNamedRangeSchema>;

/** A host comment / thread surfaced in the snapshot. */
export const DocStateCommentSchema = z
  .object({
    id: z.string(),
    author: z.string().optional(),
    text: z.string(),
    anchorHint: z.string().optional(), // short hint at what the comment is anchored to
  })
  .strict();
export type DocStateComment = z.infer<typeof DocStateCommentSchema>;

/** The ambient snapshot of the active document, injected each assist turn. */
export const DocStateSnapshotSchema = z
  .object({
    surface: SurfaceSchema,
    version: z.number().int().nonnegative(), // snapshot/schema version
    capturedAt: z.string(), // ISO 8601 capture time
    title: z.string().optional(),
    outline: z.array(DocStateOutlineEntrySchema), // headings/sections, in document order
    inventory: z.array(DocStateInventoryEntrySchema), // sheets/slides/tables/sections present
    selection: DocStateSelectionSchema.optional(),
    namedRanges: z.array(DocStateNamedRangeSchema).optional(), // Excel
    comments: z.array(DocStateCommentSchema).optional(),
    truncated: z.boolean().optional(), // true ⇒ snapshot was capped
  })
  .strict();
export type DocStateSnapshot = z.infer<typeof DocStateSnapshotSchema>;
