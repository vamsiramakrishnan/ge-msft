import type { Intent } from './intent.js';
import type { Surface } from './context.js';

/**
 * The command palette — the per-surface `/` verb list and the `@` mention kinds the input affords.
 * It is the discoverable, typed front of the same grammar the model emits (ADR-0004): each
 * {@link CommandVerb} is an `Intent` rendered as a `/verb` the user can pick, scoped by capability
 * closure (ADR-0006) so a surface never advertises a verb it cannot run.
 */

/** One `/`-verb in the palette: the intent it dispatches, its `/label`, and a one-line description. */
export interface CommandVerb {
  intent: Intent;
  label: string;
  description: string;
}

/** The kinds of `@`-mention the input affords (the grounding picker). */
export interface CommandPaletteSpec {
  surface: Surface;
  verbs: CommandVerb[];
  mentionKinds: ('document' | 'person' | 'datastore' | 'this' | 'upload')[];
}

/** The fixed `@`-mention kinds offered on every surface's input. */
const MENTION_KINDS: CommandPaletteSpec['mentionKinds'] = [
  'document',
  'person',
  'datastore',
  'this',
  'upload',
];

/** Every intent, as a palette verb. Subsets of this list are offered per surface. */
const VERB_BY_INTENT: Record<Intent, CommandVerb> = {
  assist: { intent: 'assist', label: '/assist', description: 'Grounded chat over the unit.' },
  review: {
    intent: 'review',
    label: '/review',
    description: 'Inline review pass — flag issues as findings.',
  },
  'resolve-comment': {
    intent: 'resolve-comment',
    label: '/resolve',
    description: 'Edit, reply to, and resolve a comment.',
  },
  'regen-clause': {
    intent: 'regen-clause',
    label: '/rewrite',
    description: 'Rewrite the selected clause as a tracked change.',
  },
  'draft-slides': {
    intent: 'draft-slides',
    label: '/draft',
    description: 'Generate slides from the research unit.',
  },
  synthesize: {
    intent: 'synthesize',
    label: '/synthesize',
    description: 'Synthesize the grounded sources onto the page.',
  },
  'meeting-notes': {
    intent: 'meeting-notes',
    label: '/notes',
    description: 'Live meeting notes and action items.',
  },
};

/** The intents each surface offers in its palette (before any closure narrowing). */
const VERBS_BY_SURFACE: Record<Surface, Intent[]> = {
  word: ['assist', 'review', 'resolve-comment', 'regen-clause'],
  excel: ['assist', 'review'],
  powerpoint: ['assist', 'draft-slides'],
  onenote: ['assist', 'synthesize'],
  outlook: ['assist', 'synthesize'],
  teams: ['assist', 'meeting-notes'],
};

/**
 * The command palette for a surface, optionally narrowed by capability closure (ADR-0006). The
 * `mentionKinds` are fixed; the `verbs` are the surface's intents, dropping any not in
 * `allowedIntents` when it is supplied, so the palette never offers an unreachable verb.
 */
export function commandPaletteFor(
  surface: Surface,
  allowedIntents?: Iterable<Intent>,
): CommandPaletteSpec {
  const allowed = allowedIntents ? new Set(allowedIntents) : undefined;
  const verbs = VERBS_BY_SURFACE[surface]
    .filter((intent) => (allowed ? allowed.has(intent) : true))
    .map((intent) => VERB_BY_INTENT[intent]);
  return { surface, verbs, mentionKinds: [...MENTION_KINDS] };
}
