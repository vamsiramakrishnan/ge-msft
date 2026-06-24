/**
 * The context-menu → task-pane seed contract, kept side-effect-free so the pane boot
 * (`taskpane/main.tsx`) can import it without pulling in the commands runtime's
 * `Office.actions.associate` registration (which lives in `commands.ts`).
 *
 * Security posture (see security review): the function command and the task pane are separate
 * runtimes, so the handoff must cross same-origin storage — hence `localStorage`, not
 * `sessionStorage` (which a sibling runtime can't see). To keep that channel safe:
 *   - the seed carries only an enum `kind` + a non-sensitive `hasSelection` flag — never the raw
 *     selected text (untrusted, potentially confidential) and never a free-text query;
 *   - the pane builds a FIXED query template from the kind, so a planted seed can't inject an
 *     arbitrary grounded turn or widen grounding beyond `@this`;
 *   - the consumer validates with {@link isAskSelectionSeed} and clears the key before acting.
 * The live selection is re-grounded as `@this` by the bridge at turn time, so nothing needs to
 * carry the selection text across the handoff.
 */

/** The storage key the task pane reads on boot to pick up a context-menu `assist` seed. */
export const ASK_SELECTION_SEED_KEY = 'ge:ask-selection-seed';

/** The handoff payload: a fixed kind + whether a non-empty selection existed. No content, no query. */
export interface AskSelectionSeed {
  kind: 'ask-selection';
  /** Whether the user had a non-empty selection (picks the query wording; not sensitive). */
  hasSelection: boolean;
}

/** Build the seed from the selected text — recording only whether a selection existed. */
export function buildAskSelectionSeed(selection: string): AskSelectionSeed {
  return { kind: 'ask-selection', hasSelection: selection.trim().length > 0 };
}

/** The fixed `assist` query the pane runs for the seed. Always grounds as `@this`, nothing else. */
export function askSelectionQuery(seed: AskSelectionSeed): string {
  return seed.hasSelection
    ? '@this Summarize this concisely, keeping the key facts and figures.'
    : '@this';
}

/** Narrow an untrusted parsed value to a real seed — rejects anything a foreign writer could plant. */
export function isAskSelectionSeed(value: unknown): value is AskSelectionSeed {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as { kind?: unknown }).kind === 'ask-selection' &&
    typeof (value as { hasSelection?: unknown }).hasSelection === 'boolean'
  );
}
