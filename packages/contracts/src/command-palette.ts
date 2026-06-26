import { type Intent, type CommandScope, type GroundSource } from './intent.js';
import type { Surface } from './context.js';

/**
 * The command palette — the per-surface `/` verb list, the `@` ground vocabulary the input affords,
 * and the surface-named scope options the composer renders as a segmented control. It is the
 * discoverable, typed front of the same grammar the model emits (ADR-0004): each {@link CommandVerb}
 * is an `Intent` rendered as a `/verb` the user can pick, scoped by capability closure (ADR-0006) so
 * a surface never advertises a verb it cannot run. The scope labels live here as **data**
 * (EXPERIENCE.md §4) so `web-shell` stays surface-agnostic.
 */

/** One `/`-verb in the palette: the intent it dispatches, its `/label`, and a one-line description. */
export interface CommandVerb {
  intent: Intent;
  label: string;
  description: string;
}

/** One surface-named scope choice the composer renders next to Send (the segmented control). */
export interface ScopeOption {
  scope: CommandScope;
  label: string;
}

/** The palette for a surface: its verbs, its `@`-ground kinds, and its scope options. */
export interface CommandPaletteSpec {
  surface: Surface;
  verbs: CommandVerb[];
  mentionKinds: GroundSource[];
  scopeOptions: ScopeOption[];
}

/**
 * The fixed `@`-ground kinds offered in the free-text composer. Keep this to reference kinds that
 * can resolve without an additional object picker. Addressable kinds (`document`, `person`,
 * `datastore`, `upload`) are still valid in the structured grounding model, but they require a ref
 * selected from catalog/context UI rather than a bare `@datastore` token that would be dropped.
 */
const MENTION_KINDS: GroundSource[] = ['this', 'unit'];

/**
 * Genuine label exceptions only. The default `/label` is `'/' + intent` (EXPERIENCE.md §1); this map
 * holds intents whose human label differs from the bare verb. None remain today — the override map is
 * dropped — but the seam stays so a future exception is a one-line addition, not a refactor.
 */
const LABEL_OVERRIDES: Partial<Record<Intent, string>> = {};

const DESCRIPTIONS: Record<Intent, string> = {
  ask: 'Grounded chat — ask anything about the scope.',
  summarize: 'Condense the scope.',
  explain: 'Clarify the scope in plain language.',
  rewrite: 'Apply an instruction to the scope as a reversible edit.',
  review: 'Whole-scope pass — flag issues as gated findings.',
  visualize: 'Create a chart or visual object from the scope.',
  draft: 'Generate new material from the unit.',
  notes: 'Live meeting notes and action items.',
};

/** Every intent, as a palette verb. Subsets of this list are offered per surface. */
const VERB_BY_INTENT: Record<Intent, CommandVerb> = (Object.keys(DESCRIPTIONS) as Intent[]).reduce(
  (acc, intent) => {
    acc[intent] = {
      intent,
      label: LABEL_OVERRIDES[intent] ?? `/${intent}`,
      description: DESCRIPTIONS[intent],
    };
    return acc;
  },
  {} as Record<Intent, CommandVerb>,
);

/** The intents each surface offers in its palette (before any closure narrowing). */
export const VERBS_BY_SURFACE: Record<Surface, Intent[]> = {
  word: ['ask', 'summarize', 'explain', 'rewrite', 'review'],
  excel: ['ask', 'summarize', 'explain', 'rewrite', 'review', 'visualize'],
  powerpoint: ['ask', 'summarize', 'explain', 'draft', 'review'],
  onenote: ['ask', 'summarize', 'explain', 'draft'],
  outlook: ['ask', 'summarize', 'explain', 'draft'],
  teams: ['ask', 'summarize', 'notes'],
};

/**
 * The surface-named scope options the composer renders (EXPERIENCE.md §4 — labels are data, so the
 * surface-agnostic `Composer` reads them rather than hard-coding "range | slide | thread"). The
 * first entry is the per-surface default.
 */
const SCOPE_OPTIONS_BY_SURFACE: Record<Surface, ScopeOption[]> = {
  word: [
    { scope: { kind: 'selection' }, label: 'Selection' },
    { scope: { kind: 'document' }, label: 'Whole document' },
    { scope: { kind: 'section' }, label: 'This section' },
  ],
  excel: [
    { scope: { kind: 'selection' }, label: 'Selection' },
    { scope: { kind: 'document' }, label: 'Sheet' },
    { scope: { kind: 'range' }, label: 'Range' },
  ],
  powerpoint: [
    { scope: { kind: 'this-item' }, label: 'This slide' },
    { scope: { kind: 'document' }, label: 'Deck' },
  ],
  onenote: [{ scope: { kind: 'document' }, label: 'Page' }],
  outlook: [
    { scope: { kind: 'this-item' }, label: 'This message' },
    { scope: { kind: 'document' }, label: 'Whole thread' },
  ],
  teams: [
    { scope: { kind: 'document' }, label: 'Transcript' },
    { scope: { kind: 'range', ref: 'last-5-min' }, label: 'Last 5 min' },
  ],
};

/**
 * The command palette for a surface, optionally narrowed by capability closure (ADR-0006). The
 * `mentionKinds` and `scopeOptions` are fixed per surface; the `verbs` are the surface's intents,
 * dropping any not in `allowedIntents` when it is supplied, so the palette never offers an
 * unreachable verb.
 */
export function commandPaletteFor(
  surface: Surface,
  allowedIntents?: Iterable<Intent>,
): CommandPaletteSpec {
  const allowed = allowedIntents ? new Set(allowedIntents) : undefined;
  const verbs = VERBS_BY_SURFACE[surface]
    .filter((intent) => (allowed ? allowed.has(intent) : true))
    .map((intent) => VERB_BY_INTENT[intent]);
  return {
    surface,
    verbs,
    mentionKinds: [...MENTION_KINDS],
    scopeOptions: SCOPE_OPTIONS_BY_SURFACE[surface].map((o) => ({ ...o, scope: { ...o.scope } })),
  };
}
