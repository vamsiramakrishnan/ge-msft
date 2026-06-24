import type { Surface, Intent, CommandScope } from '@ge/contracts';

/**
 * The context-menu → task-pane seed contract, kept side-effect-free so the pane boot
 * (`taskpane/main.tsx`) can import it without pulling in the commands runtime's
 * `Office.actions.associate` registration (which lives in `commands.ts`).
 *
 * Security posture (see security review): the function command and the task pane are separate
 * runtimes, so the handoff must cross same-origin storage — hence `localStorage`, not
 * `sessionStorage` (which a sibling runtime can't see). To keep that channel safe:
 *   - the seed carries only a typed `{intent, scope}` + a non-sensitive `hasSelection` flag —
 *     never the raw selected text (untrusted, potentially confidential) and never a free-text query;
 *   - the pane builds a FIXED query template from the typed fields, so a planted seed can't inject
 *     an arbitrary grounded turn or widen grounding beyond `@this`;
 *   - the consumer validates with {@link isAskSelectionSeed} (which pins the schema `version`) and
 *     clears the key before acting; a `nonce` + `ts` + TTL bound a stale/replayed seed.
 * The live selection is re-grounded as `@this` by the bridge at turn time, so nothing needs to
 * carry the selection text across the handoff.
 */

/** The seed schema version — bumped on any shape change; the consumer rejects a mismatch. */
export const ASK_SELECTION_SEED_VERSION = 1;

/** How long a stashed seed stays valid (ms). A boot past this drops it as stale rather than firing. */
export const ASK_SELECTION_SEED_TTL_MS = 60_000;

/** The base storage key — namespaced per surface so two open hosts can't read each other's seed. */
const ASK_SELECTION_SEED_BASE = 'ge:ask-selection-seed';

/** The per-surface storage key the task pane reads on boot to pick up a context-menu seed. */
export function askSelectionSeedKey(surface: Surface): string {
  return `${ASK_SELECTION_SEED_BASE}:${surface}`;
}

/**
 * The handoff payload: a pinned `version`, the typed `{intent, scope}` the pane runs, a `nonce` +
 * `ts` for replay/staleness bounding, and whether a non-empty selection existed. No content, no
 * free-text query.
 */
export interface AskSelectionSeed {
  kind: 'ask-selection';
  version: number;
  /** The verb the pane runs — the right-click entry is a grounded `ask` over `@this`. */
  intent: Intent;
  /** WHERE the verb acts — right-click hard-binds the live `selection`. */
  scope: CommandScope;
  /** Whether the user had a non-empty selection (picks the query wording; not sensitive). */
  hasSelection: boolean;
  /** A random nonce — distinguishes two seeds and lets a consumer dedupe a replay. */
  nonce: string;
  /** Epoch-ms the seed was written; the consumer drops it once older than the TTL. */
  ts: number;
}

/** A small random nonce — `crypto` when available, falling back to a time+random string. */
function makeNonce(): string {
  const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
  if (c?.randomUUID) return c.randomUUID();
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

/** Build the seed from the selected text — recording only whether a selection existed, plus metadata. */
export function buildAskSelectionSeed(
  selection: string,
  now: number = Date.now(),
): AskSelectionSeed {
  return {
    kind: 'ask-selection',
    version: ASK_SELECTION_SEED_VERSION,
    intent: 'ask',
    scope: { kind: 'selection' },
    hasSelection: selection.trim().length > 0,
    nonce: makeNonce(),
    ts: now,
  };
}

/** The fixed `ask` query the pane runs for the seed. Always grounds as `@this`, nothing else. */
export function askSelectionQuery(seed: AskSelectionSeed): string {
  return seed.hasSelection
    ? '@this Summarize this concisely, keeping the key facts and figures.'
    : '@this';
}

/** True when the seed is within its TTL window (not stale/replayed across a long-idle pane). */
export function isAskSelectionSeedFresh(seed: AskSelectionSeed, now: number = Date.now()): boolean {
  return now - seed.ts <= ASK_SELECTION_SEED_TTL_MS && now >= seed.ts;
}

/**
 * Narrow an untrusted parsed value to a real seed — rejects anything a foreign writer could plant,
 * including a seed written by an older/newer shell (the `version` must match exactly).
 */
export function isAskSelectionSeed(value: unknown): value is AskSelectionSeed {
  if (typeof value !== 'object' || value === null) return false;
  const v = value as Record<string, unknown>;
  return (
    v.kind === 'ask-selection' &&
    v.version === ASK_SELECTION_SEED_VERSION &&
    v.intent === 'ask' &&
    typeof v.scope === 'object' &&
    v.scope !== null &&
    (v.scope as { kind?: unknown }).kind === 'selection' &&
    typeof v.hasSelection === 'boolean' &&
    typeof v.nonce === 'string' &&
    typeof v.ts === 'number'
  );
}
