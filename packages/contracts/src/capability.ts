import { z } from 'zod';
import { ChangeIdSchema } from './brand.js';
import { ContextKindSchema, SurfaceSchema } from './context.js';
import { SourceRefSchema } from './finding.js';
import { ProvenancePayloadSchema } from './provenance.js';

/**
 * The capability foundation (part 2 of 2): **actuation**.
 *
 * Everything a surface can *write* is advertised as an `Actuation` and invoked via an
 * `ActuationRequest`. Writes are reversible and provenanced by contract: each carries a
 * client-generated `changeId` (idempotent re-apply) and an optional `ProvenancePayload`
 * stamped into the host's durable metadata. Experiences/agents compose these; they never
 * touch Office.js directly. See docs/ADR-0002-capability-model.md.
 */

export const ActuationKindSchema = z.enum([
  'insert-text', // insert plain text at the cursor/selection
  'replace-selection', // replace the current selection
  'tracked-change', // Word: insert/replace as a tracked change
  'insert-ooxml', // Word/PPT: insert rich OOXML
  'fill-content-control', // Word: populate a named content control
  'add-comment', // Word/Excel: add a new content/cell-anchored comment (ADR-0004 `comment` verb)
  'comment-reply', // Word/Excel/PPT: reply to (and optionally resolve) a comment
  'write-cells', // Excel: write values/formulas to a range
  'format-cells', // Excel: apply formatting (bold/fill/numberFormat) to a range (ADR-0004 `format`)
  'set-entity-card', // Excel: attach a linked-entity card to a cell
  'insert-slide', // PowerPoint: add a slide
  'set-speaker-notes', // PowerPoint: set a slide's speaker notes
  'append-page', // OneNote: add a synthesized page
  'create-mail', // Outlook: open a draft message
  'reply-mail', // Outlook: open a grounded reply
  'create-event', // Outlook/Graph: create a calendar item
  'create-task', // Planner/To Do: create a task
  'post-card', // Teams: post an Adaptive Card (notes/action items)
  'post-message', // Teams: stage a reviewable chat post / Adaptive Card (reversible)
]);
export type ActuationKind = z.infer<typeof ActuationKindSchema>;

/** What a surface advertises it can do (drives the UI's available actions). */
export const ActuationSchema = z.object({
  kind: ActuationKindSchema,
  surface: SurfaceSchema,
  title: z.string(), // verb shown to the user ("Insert as tracked change")
  description: z.string().optional(),
  reversible: z.boolean(), // false ⇒ the UI must warn before invoking
  appliesTo: z.array(ContextKindSchema).optional(), // context kinds this can target
});
export type Actuation = z.infer<typeof ActuationSchema>;

/** Parameters for an actuation. Open by design — agents fill what a kind needs. */
export const ActuationParamsSchema = z.object({
  text: z.string().optional(),
  ooxml: z.string().optional(),
  html: z.string().optional(),
  /** A content-anchored target (matchText/contextHint) or an explicit host id. */
  target: z
    .object({
      matchText: z.string().optional(),
      contextHint: z.string().optional(),
      contentControlId: z.string().optional(),
      commentId: z.string().optional(),
      range: z.string().optional(),
      slideIndex: z.number().optional(),
    })
    .optional(),
  cells: z.array(z.array(z.string())).optional(), // write-cells
  /** format-cells: host-native formatting applied to `target.range` (ADR-0004 `format` verb). */
  format: z
    .object({
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      fill: z.string().optional(), // background color, e.g. "#FFF2CC"
      numberFormat: z.string().optional(), // e.g. "$#,##0.00"
    })
    .optional(),
  slide: z
    .object({ title: z.string(), bullets: z.array(z.string()), notes: z.string().optional() })
    .optional(),
  mail: z
    .object({
      to: z.array(z.string()).optional(),
      subject: z.string().optional(),
      body: z.string().optional(),
    })
    .optional(),
  resolveComment: z.boolean().optional(),
  sources: z.array(SourceRefSchema).optional(),
});
export type ActuationParams = z.infer<typeof ActuationParamsSchema>;

export const ActuationRequestSchema = z.object({
  changeId: ChangeIdSchema, // client-generated; makes the write idempotent
  kind: ActuationKindSchema,
  surface: SurfaceSchema,
  params: ActuationParamsSchema,
  provenance: ProvenancePayloadSchema.optional(),
});
export type ActuationRequest = z.infer<typeof ActuationRequestSchema>;

export const ActuationResultSchema = z.object({
  ok: z.boolean(),
  changeId: ChangeIdSchema,
  kind: ActuationKindSchema,
  location: z.string().optional(), // where it landed (range, slide #, comment id, draft id)
  degraded: z.boolean().optional(), // e.g. anchor drifted → applied as a panel item
  error: z.object({ code: z.string(), message: z.string() }).optional(),
});
export type ActuationResult = z.infer<typeof ActuationResultSchema>;

/** A surface's full capability advertisement: what it can read and what it can write. */
export const CapabilityManifestSchema = z.object({
  surface: SurfaceSchema,
  contextKinds: z.array(ContextKindSchema),
  actuations: z.array(ActuationSchema),
});
export type CapabilityManifest = z.infer<typeof CapabilityManifestSchema>;
