import { z } from 'zod';
import { ChangeIdSchema } from './brand.js';
import { ContextKindSchema, SurfaceSchema, type Surface } from './context.js';
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
  'create-table', // Excel: promote a range to a native Table (ADR-0007 `table` verb)
  'insert-chart', // Excel: add a chart over a source range (ADR-0007 `chart` verb)
  'format-conditional', // Excel: add a conditional-format rule to a range (ADR-0007 `cf` verb)
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
  /** create-table (ADR-0007): promote `range` to a native Table. */
  table: z
    .object({
      range: z.string(),
      hasHeaders: z.boolean().default(true),
      name: z.string().optional(), // table name; the bridge mints one if absent
    })
    .optional(),
  /** insert-chart (ADR-0007): a chart over `sourceRange`. */
  chart: z
    .object({
      chartType: z.enum(['column', 'bar', 'line', 'pie', 'scatter', 'area']),
      sourceRange: z.string(),
      seriesBy: z.enum(['rows', 'columns', 'auto']).default('auto'),
      title: z.string().optional(),
    })
    .optional(),
  /** format-conditional (ADR-0007): one conditional-format rule applied to `range`. */
  conditional: z
    .object({
      range: z.string(),
      rule: z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('cellValue'),
          operator: z.enum(['gt', 'lt', 'ge', 'le', 'eq', 'ne', 'between']),
          value: z.string(),
          value2: z.string().optional(), // upper bound for `between`
          fill: z.string().optional(), // highlight color, e.g. "#C6EFCE"
        }),
        z.object({ kind: z.literal('dataBar') }),
        z.object({ kind: z.literal('colorScale') }),
        z.object({
          kind: z.literal('top'),
          rank: z.number(), // top/bottom N
          bottom: z.boolean().default(false),
          fill: z.string().optional(),
        }),
      ]),
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

/**
 * The recorded INVERSE of an actuation (ADR-0007). Reversibility is an explicit, recorded operation
 * rather than an implicit property of a tracked change: when a bridge lands a write it reports HOW to
 * undo it. Pure additions (table/chart/pivot) carry a `delete-object` descriptor keyed by the object
 * name the host minted; in-place mutations (conditional formatting, a grid spilled over existing data)
 * carry the prior state needed to restore it. The descriptor is persisted alongside provenance so an
 * undo is auditable and does not depend on host-session undo state.
 */
export const InverseDescriptorSchema = z.discriminatedUnion('op', [
  z.object({
    op: z.literal('delete-object'),
    objectType: z.enum(['table', 'chart', 'pivot']),
    name: z.string(), // the host object name to delete
  }),
  z.object({
    op: z.literal('restore-values'),
    range: z.string(),
    values: z.array(z.array(z.string())), // prior cell values to write back
  }),
  z.object({
    op: z.literal('clear-conditional'),
    range: z.string(),
    ruleOrdinal: z.number(), // index of the added rule within the range's CF collection
  }),
]);
export type InverseDescriptor = z.infer<typeof InverseDescriptorSchema>;

export const ActuationResultSchema = z.object({
  ok: z.boolean(),
  changeId: ChangeIdSchema,
  kind: ActuationKindSchema,
  location: z.string().optional(), // where it landed (range, slide #, comment id, draft id)
  // How to reverse this write (ADR-0007). Populated by the bridge at apply-time once it knows the
  // minted object name / prior state; persisted with provenance so undo is recorded, not implicit.
  inverse: InverseDescriptorSchema.optional(),
  degraded: z.boolean().optional(), // e.g. anchor drifted → applied as a panel item
  // Observability: the reversible write LANDED but its durable provenance could not be persisted
  // (host metadata write failed / unavailable). The change is real but unprovenanced — surface it so
  // the audit trail and the user know, rather than silently dropping the trace. Absent ⇒ recorded.
  provenanceDropped: z.boolean().optional(),
  // Observability: the write LANDED carrying NO provenance payload at all (the turn produced no
  // provenance to stamp). Distinct from `provenanceDropped` (had a record, failed to persist) — this
  // is an unattributed write, surfaced so it is never mistaken for an attributed one.
  provenanceMissing: z.boolean().optional(),
  error: z.object({ code: z.string(), message: z.string() }).optional(),
});
export type ActuationResult = z.infer<typeof ActuationResultSchema>;

/**
 * The read verbs a surface can serve (ADR-0006 capability closure). The CLI grammar advertises a
 * read verb only when it appears here, and conformance tests require a matching bridge read port —
 * so a surface can never advertise a read it cannot perform. Optional for back-compat; an absent
 * value means "no addressable reads declared".
 */
export const ReadVerbSchema = z.enum(['outline', 'read', 'search']);

/** A surface's full capability advertisement: what it can read and what it can write. */
export const CapabilityManifestSchema = z.object({
  surface: SurfaceSchema,
  contextKinds: z.array(ContextKindSchema),
  actuations: z.array(ActuationSchema),
  /** Read verbs this surface serves (ADR-0006); the grammar scopes `outline`/`read`/`search` to it. */
  reads: z.array(ReadVerbSchema).optional(),
});
export type CapabilityManifest = z.infer<typeof CapabilityManifestSchema>;

/**
 * ADR-0006 — the `Capability` descriptor: the forward source of truth for a single capability.
 *
 * One descriptor names a capability, locates it on a surface, and classifies it on the
 * pure/effect split that ADR-0005 makes load-bearing:
 *   - `read`   — a Layer-B host read (produces a value; never gated).
 *   - `pure`   — a pure transform over values (composes freely; never gated).
 *   - `effect` — an actuation terminal (consumes values, produces a gated `Effect`).
 *
 * The intent (ADR-0006) is that the manifest, the verb→kind map, and dispatch are eventually
 * *derived* from a registry of these descriptors for new capabilities. This is a typed scaffold:
 * no migration is required this wave — the {@link checkCapabilityClosure} conformance gate is what
 * makes that incremental migration safe (drift can't silently return while descriptors and the
 * hand-written manifest/map coexist). `signature` and `gatePolicy` are deliberately open for now;
 * later waves narrow them as the registry lands.
 */
export interface Capability {
  /** Stable capability name (also the CLI/skill identifier, e.g. `reply`, `write-cells`). */
  name: string;
  /** The surface this capability is defined on. */
  surface: Surface;
  /** The pure/effect classification (the ADR-0005 composition + safety boundary). */
  kind: 'read' | 'pure' | 'effect';
  /** A forward type signature for the value layer (open this wave; narrowed later). */
  signature?: unknown;
  /** A forward gate-policy hook for `effect` capabilities (open this wave; narrowed later). */
  gatePolicy?: unknown;
}
