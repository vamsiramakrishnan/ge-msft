import { z } from 'zod';
import { ActuationKindSchema, ActuationRequestSchema, type ActuationKind } from './capability.js';

/**
 * ADR-0008 §7 — the dependency-aware effect plan (the DAG the runtime compiles a linear program into).
 *
 * Algebraic source is linear because it's easy to write and read; execution is NOT a naïve list. The
 * compiler infers a DAG from variable use, overlapping read/write resources, derived ranges, and
 * effect ordering — the model never writes `depends-on` syntax. This is a **dependency-aware effect
 * plan / saga with bounded compensation**, NOT an atomic transaction (a bridge may lack a reliable
 * inverse — `reversible:false` kinds exist). On a node's failure, its dependents are skipped
 * (`prerequisite_failed`); independent effects may still run if policy permits.
 */

/** A typed handle to a resource an effect reads or writes (for dependency inference). */
export const ResourceRefSchema = z.object({
  // `range` (A1, normalized) · `object` (a minted table/chart/shape name) · `anchor` (content match) ·
  // `comment` (a comment id) · `draft` (the open compose draft) · `estate` (a Graph-side target).
  kind: z.enum(['range', 'object', 'anchor', 'comment', 'draft', 'estate']),
  id: z.string(),
});
export type ResourceRef = z.infer<typeof ResourceRefSchema>;

/**
 * The approval authority an effect carries — distinct classes must not silently share one approval
 * (ADR-0008 §break-boundaries). `irreversible` is the strongest gate.
 */
export const ApprovalClassSchema = z.enum(['in-document', 'external', 'estate', 'irreversible']);
export type ApprovalClass = z.infer<typeof ApprovalClassSchema>;

export const FailurePolicySchema = z.enum(['stop-dependents', 'continue-independent']);
export type FailurePolicy = z.infer<typeof FailurePolicySchema>;

/** A pure node — a read or a transform; evaluated before approval, never gated. */
export const PurePlanNodeSchema = z.object({
  kind: z.literal('pure'),
  id: z.string(),
  dependsOn: z.array(z.string()),
  operator: z.string(), // 'read' | 'filter' | 'sort' | … | a `let` binding name
  outputType: z.string(), // a ValueType ('Table'/'Number'/…)
});
export type PurePlanNode = z.infer<typeof PurePlanNodeSchema>;

/** An effect node — a gated host mutation, with the resources + policy the gate/saga need. */
export const EffectPlanNodeSchema = z.object({
  kind: z.literal('effect'),
  id: z.string(),
  dependsOn: z.array(z.string()),
  request: ActuationRequestSchema,
  reads: z.array(ResourceRefSchema),
  writes: z.array(ResourceRefSchema),
  approvalClass: ApprovalClassSchema,
  reversible: z.boolean(),
  idempotencyKey: z.string(),
  failurePolicy: FailurePolicySchema,
});
export type EffectPlanNode = z.infer<typeof EffectPlanNodeSchema>;

export const PlanNodeSchema = z.discriminatedUnion('kind', [
  PurePlanNodeSchema,
  EffectPlanNodeSchema,
]);
export type PlanNode = z.infer<typeof PlanNodeSchema>;

/**
 * Actuation kinds with NO reliable inverse — the saga cannot compensate them, so they are
 * `reversible:false` and approval-class `irreversible` (warn hard; never silently grouped). Mirrors
 * the `not-reversible` inverse marker + the catalogue's "hard" rows.
 */
export const IRREVERSIBLE_KINDS: ReadonlySet<ActuationKind> = new Set([
  'insert-text',
  'insert-ooxml',
  'resolve-revisions',
  'send-activity-notification',
  'delete-message',
  'create-mail-rule',
]);

/** Kinds that send to an external recipient / leave the open document (vs an in-document change). */
export const EXTERNAL_KINDS: ReadonlySet<ActuationKind> = new Set([
  'reply-mail',
  'create-mail',
  'post-message',
  'post-card',
]);

/** Plane-B estate kinds (Microsoft Graph) — a distinct, higher approval authority. */
export const ESTATE_KINDS: ReadonlySet<ActuationKind> = new Set([
  'create-event',
  'create-task',
  'move-message',
  'copy-message',
  'categorize-message',
  'flag-message',
  'delete-message',
  'create-mail-rule',
  'post-chat-message',
  'post-channel-message',
  'reply-channel-message',
  'update-message',
  'set-reaction',
  'create-online-meeting',
  'send-activity-notification',
  'graph-create-page',
  'graph-patch-page',
  'graph-create-section',
]);

/** The approval authority for a kind (ADR-0008 §break-boundaries): irreversible > estate > external > in-document. */
export function approvalClassOf(kind: ActuationKind): ApprovalClass {
  if (IRREVERSIBLE_KINDS.has(kind)) return 'irreversible';
  if (ESTATE_KINDS.has(kind)) return 'estate';
  if (EXTERNAL_KINDS.has(kind)) return 'external';
  return 'in-document';
}

/** Whether a kind has a reliable inverse (the saga can compensate it). */
export function isReversibleKind(kind: ActuationKind): boolean {
  return !IRREVERSIBLE_KINDS.has(kind);
}

// Re-export so consumers importing from plan-graph get the kind schema without a second import.
export { ActuationKindSchema };
