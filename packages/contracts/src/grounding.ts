import { z } from 'zod';
import { type GroundSource } from './intent.js';

/**
 * GROUNDING SELECTION (the typed `@`-mention value).
 *
 * Finding #2 from the review: the composer parses `@`-mentions but throws the parse away — `@this`
 * ends up as raw prompt *text*, not a typed grounding reference, so it never scopes the turn. This
 * module is the missing typed value: a {@link GroundingSelection} is the addressable thing the user
 * picked from the `@`-picker, carrying both the {@link GroundSource} *kind* and the id needed to
 * resolve it onto a `streamAssist` request (done by `resolveGrounding` in `@ge/gemini-client`).
 *
 * Relationship to {@link GroundSource}: `GroundSource` is the vocabulary of *kinds* the picker
 * affords (`this`/`unit`/`document`/`person`/`datastore`/`upload`). `GroundingSelection` is one
 * *concrete* pick: it keeps the same kinds (renamed for clarity at the value level — `this` →
 * `current-context`, `datastore` → `data-store`) and adds the addressable id where the kind needs
 * one. The reference-kinds (`current-context`, `unit`) address the live scope/unit the bridge
 * already attaches and so carry no id; the rest carry the id that makes them resolvable.
 */
export const GroundingSelectionSchema = z.discriminatedUnion('kind', [
  /** The live scope ("this {selection|range|slide|thread}") — the context the bridge attaches. */
  z.object({ kind: z.literal('current-context') }),
  /** The research unit (notebook + federated sources + working document). */
  z.object({ kind: z.literal('unit') }),
  /** A named document already indexed in a connected data store, by its VAIS document name. */
  z.object({ kind: z.literal('document'), id: z.string().min(1) }),
  /** A person/contact reference, by display name or person id. */
  z.object({ kind: z.literal('person'), id: z.string().min(1) }),
  /** A whole connected data store, by its full resource name. */
  z.object({ kind: z.literal('data-store'), id: z.string().min(1) }),
  /** An ad-hoc uploaded source, by the session context `fileId` it was stored as. */
  z.object({ kind: z.literal('upload'), fileId: z.string().min(1) }),
]);
export type GroundingSelection = z.infer<typeof GroundingSelectionSchema>;

/** The `kind` of a {@link GroundingSelection} — the value-level spelling of {@link GroundSource}. */
export type GroundingSelectionKind = GroundingSelection['kind'];

/**
 * The `@`-picker emits a {@link GroundSource} kind; this maps that vocabulary onto the value-level
 * {@link GroundingSelectionKind}. Kept total and explicit so adding a `GroundSource` is a compile
 * error here until the value spelling is decided, rather than a silently-dropped mention.
 */
export const GROUND_SOURCE_TO_SELECTION_KIND: Record<GroundSource, GroundingSelectionKind> = {
  this: 'current-context',
  unit: 'unit',
  document: 'document',
  person: 'person',
  datastore: 'data-store',
  upload: 'upload',
};
