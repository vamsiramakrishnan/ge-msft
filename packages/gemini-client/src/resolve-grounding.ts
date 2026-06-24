import type { GroundingSelection, GroundingSelectionKind } from '@ge/contracts';
import type { QueryPart } from './session-context.js';
import type { DataStoreSpec } from './search.js';

/**
 * Resolve typed `@`-mentions ({@link GroundingSelection}) into the `streamAssist` request fields they
 * address. This is the missing half of review Finding #2: the composer parses `@`-mentions but the
 * typed value was discarded, so `@this` ended up as raw prompt *text* rather than a grounding ref.
 * Here each selection lands in the right request bucket — and crucially, **no selection is ever
 * turned into free-text prompt content**: an unresolvable pick is dropped with a structured note,
 * never inlined and never thrown.
 *
 * The mapping (per docs/api/discoveryengine/context-mechanisms.md + streamAssist.md):
 * - `current-context` / `unit` → the `query.parts[]` the bridge already attaches (supplied via ctx).
 * - `document` → `query.parts[].documentReference` (a VAIS document by name).
 * - `person`   → `query.parts[].personReference`.
 * - `data-store` → `toolsSpec.vertexAiSearchSpec.dataStoreSpecs[]`.
 * - `upload`   → top-level `fileIds[]` (a session context file).
 *
 * Pure and total: it never mutates its inputs and never throws.
 */

/** Why a selection could not be resolved — surfaced (not swallowed) so the UI can explain the drop. */
export interface DroppedGroundingNote {
  kind: GroundingSelectionKind;
  reason: string;
}

/**
 * What the bridge/session already knows, so the reference-kinds resolve to the *same* query parts the
 * live scope and unit are attached as. The resolver does not re-derive these from host content — it
 * only re-addresses what the bridge supplies, keeping host content as data the bridge framed.
 */
export interface GroundingResolveContext {
  /** The `query.parts[]` the bridge attaches for the live scope (`@this`). */
  contextParts?: QueryPart[];
  /** The `query.parts[]` the bridge attaches for the research unit (`@unit`). */
  unitParts?: QueryPart[];
}

/**
 * The partial `streamAssist` request a set of grounding selections contributes. The wiring agent
 * merges these into the request it builds: `queryParts` extend `query.parts[]`, `dataStoreSpecs`
 * extend `toolsSpec.vertexAiSearchSpec.dataStoreSpecs`, `fileIds` extend the top-level `fileIds`.
 * Any field is omitted when it has no entries. `notes` lists selections that were dropped.
 */
export interface ResolvedGrounding {
  queryParts?: QueryPart[];
  dataStoreSpecs?: DataStoreSpec[];
  fileIds?: string[];
  notes?: DroppedGroundingNote[];
}

export function resolveGrounding(
  selections: readonly GroundingSelection[],
  ctx: GroundingResolveContext,
): ResolvedGrounding {
  const queryParts: QueryPart[] = [];
  const dataStoreSpecs: DataStoreSpec[] = [];
  const fileIds: string[] = [];
  const notes: DroppedGroundingNote[] = [];

  for (const sel of selections) {
    switch (sel.kind) {
      case 'current-context': {
        const parts = ctx.contextParts ?? [];
        if (parts.length === 0) {
          notes.push({
            kind: sel.kind,
            reason: 'No live context is attached to this turn.',
          });
          break;
        }
        queryParts.push(...parts);
        break;
      }
      case 'unit': {
        const parts = ctx.unitParts ?? [];
        if (parts.length === 0) {
          notes.push({
            kind: sel.kind,
            reason: 'No research unit is attached to this turn.',
          });
          break;
        }
        queryParts.push(...parts);
        break;
      }
      case 'document': {
        queryParts.push({ documentReference: { documentName: sel.id } });
        break;
      }
      case 'person': {
        queryParts.push({ personReference: { displayName: sel.id } });
        break;
      }
      case 'data-store': {
        dataStoreSpecs.push({ dataStore: sel.id });
        break;
      }
      case 'upload': {
        fileIds.push(sel.fileId);
        break;
      }
    }
  }

  return {
    ...(queryParts.length > 0 ? { queryParts } : {}),
    ...(dataStoreSpecs.length > 0 ? { dataStoreSpecs } : {}),
    ...(fileIds.length > 0 ? { fileIds } : {}),
    ...(notes.length > 0 ? { notes } : {}),
  };
}
