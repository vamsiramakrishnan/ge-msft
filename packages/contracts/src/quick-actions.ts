import { z } from 'zod';
import { IntentSchema, type Intent } from './intent.js';
import { SurfaceSchema, type Surface } from './context.js';

/**
 * Quick actions — the curated, one-tap verb catalog the panel and the right-click context menu
 * offer per surface. Each entry seeds the assistant with a templated free-text `prompt`, a default
 * set of `@`-grounding sources, the `intent` it dispatches as, and the shape of its `output` (a
 * chat answer, an inline annotation pass, or a write-back). It is the human-friendly front of the
 * same grammar the model emits (ADR-0004) and is scoped by capability closure (ADR-0006) — an
 * action whose intent the surface cannot run is never offered.
 *
 * `ground` defaults: read-only summarize/explain actions ground on `['this']` (the live selection /
 * open item); actions that lean on the research unit ground on `['unit']` (the notebook + federated
 * sources + working document).
 */

/** The output shape an action produces: grounded chat, an inline annotation pass, or a write-back. */
export const QuickActionOutputSchema = z.enum(['chat', 'annotation', 'write']);
export type QuickActionOutput = z.infer<typeof QuickActionOutputSchema>;

export const QuickActionSchema = z.object({
  id: z.string(),
  label: z.string(),
  surfaces: z.array(SurfaceSchema).nonempty(),
  intent: IntentSchema,
  prompt: z.string(), // templated free-text seed
  ground: z.array(z.string()).default([]), // default @-sources, e.g. ['this']
  output: QuickActionOutputSchema,
  contextMenu: z.boolean().default(false), // also offered on right-click
});
export type QuickAction = z.infer<typeof QuickActionSchema>;

/** All six surfaces — the `surfaces` value for a universal action. */
const ALL_SURFACES: [Surface, ...Surface[]] = [
  'word',
  'excel',
  'powerpoint',
  'onenote',
  'outlook',
  'teams',
];

/**
 * The quick-action catalog. Universal actions head the list; the per-surface blocks follow in the
 * surface order above. Every entry parses {@link QuickActionSchema}; the defaults (`ground: []`,
 * `contextMenu: false`) are written explicitly so the literal stays self-describing.
 */
export const QUICK_ACTIONS: QuickAction[] = [
  // ── Universal (all six surfaces) ───────────────────────────────────────────
  {
    id: 'summarize-this',
    label: 'Summarize this',
    surfaces: ALL_SURFACES,
    intent: 'assist',
    prompt: 'Summarize this concisely, keeping the key facts and figures.',
    ground: ['this'],
    output: 'chat',
    contextMenu: true,
  },
  {
    id: 'key-points',
    label: 'Key points & action items',
    surfaces: ALL_SURFACES,
    intent: 'assist',
    prompt: 'List the key points and any action items, with owners where stated.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'explain',
    label: 'Explain / clarify',
    surfaces: ALL_SURFACES,
    intent: 'assist',
    prompt: 'Explain this in plain language and clarify anything ambiguous.',
    ground: ['this'],
    output: 'chat',
    contextMenu: true,
  },
  {
    id: 'find-risks',
    label: 'Find risks & gaps',
    surfaces: ALL_SURFACES,
    intent: 'assist',
    prompt: 'Identify the risks, gaps, and open questions in this.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },

  // ── Word ───────────────────────────────────────────────────────────────────
  {
    id: 'review-policy',
    label: 'Review against policy',
    surfaces: ['word'],
    intent: 'review',
    prompt: 'Review this document against the grounded policy and flag every breach as a finding.',
    ground: ['unit'],
    output: 'annotation',
    contextMenu: false,
  },
  {
    id: 'find-unsupported',
    label: 'Find unsupported claims',
    surfaces: ['word'],
    intent: 'review',
    prompt: 'Flag claims that are unsupported by the grounded sources.',
    ground: ['unit'],
    output: 'annotation',
    contextMenu: false,
  },
  {
    id: 'tighten',
    label: 'Tighten / rewrite selection',
    surfaces: ['word'],
    intent: 'regen-clause',
    prompt: 'Tighten and rewrite the selected text as a tracked change, preserving its meaning.',
    ground: ['this'],
    output: 'write',
    contextMenu: true,
  },
  {
    id: 'exec-summary',
    label: 'Draft an executive summary',
    surfaces: ['word'],
    intent: 'assist',
    prompt: 'Draft a concise executive summary of this document.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'resolve-comment',
    label: 'Resolve this comment',
    surfaces: ['word'],
    intent: 'resolve-comment',
    prompt: 'Resolve this comment: edit the anchored text, reply to the thread, and resolve it.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },

  // ── Excel ────────────────────────────────────────────────────────────────────
  {
    id: 'summarize-range',
    label: 'Summarize this range',
    surfaces: ['excel'],
    intent: 'assist',
    prompt: 'Summarize the selected range, calling out totals, trends, and outliers.',
    ground: ['this'],
    output: 'chat',
    contextMenu: true,
  },
  {
    id: 'explain-formula',
    label: 'Explain this formula',
    surfaces: ['excel'],
    intent: 'assist',
    prompt: 'Explain what the formula in the selected cell does, step by step.',
    ground: ['this'],
    output: 'chat',
    contextMenu: true,
  },
  {
    id: 'risk-column',
    label: 'Add a risk/summary column',
    surfaces: ['excel'],
    intent: 'assist',
    prompt: 'Add a risk or summary column derived from the selected range.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },
  {
    id: 'find-anomalies',
    label: 'Find anomalies / outliers',
    surfaces: ['excel'],
    intent: 'review',
    prompt: 'Find anomalies and outliers in the selected range and comment on each.',
    ground: ['this'],
    output: 'annotation',
    contextMenu: false,
  },
  {
    id: 'write-formula',
    label: 'Write a formula for…',
    surfaces: ['excel'],
    intent: 'assist',
    prompt: 'Write a formula for {{describe what to compute}} and place it in the target cell.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },

  // ── PowerPoint ───────────────────────────────────────────────────────────────
  {
    id: 'draft-section',
    label: 'Draft a section from the unit',
    surfaces: ['powerpoint'],
    intent: 'draft-slides',
    prompt: 'Draft a slide section from the research unit on {{topic}}.',
    ground: ['unit'],
    output: 'write',
    contextMenu: false,
  },
  {
    id: 'speaker-notes',
    label: 'Generate speaker notes',
    surfaces: ['powerpoint'],
    intent: 'assist',
    prompt: 'Generate speaker notes for the current slide.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },
  {
    id: 'summarize-deck',
    label: 'Summarize the deck',
    surfaces: ['powerpoint'],
    intent: 'assist',
    prompt: 'Summarize the narrative and key takeaways of this deck.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'redesign',
    label: 'Suggest a redesign',
    surfaces: ['powerpoint'],
    intent: 'assist',
    prompt: 'Suggest a redesign for the current slide to make it clearer and tighter.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },

  // ── OneNote ──────────────────────────────────────────────────────────────────
  {
    id: 'synthesize-page',
    label: 'Summarize sources onto the page',
    surfaces: ['onenote'],
    intent: 'synthesize',
    prompt: 'Synthesize the grounded sources into a cited summary on this page.',
    ground: ['unit'],
    output: 'write',
    contextMenu: false,
  },
  {
    id: 'audio-overview',
    label: 'Make an audio overview',
    surfaces: ['onenote'],
    intent: 'assist',
    prompt: 'Produce an audio-overview script of the grounded sources.',
    ground: ['unit'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'discover-sources',
    label: 'Discover related sources',
    surfaces: ['onenote'],
    intent: 'assist',
    prompt: 'Discover sources related to this page from the connected data stores.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },

  // ── Outlook ──────────────────────────────────────────────────────────────────
  {
    id: 'summarize-email',
    label: 'Summarize this email / thread',
    surfaces: ['outlook'],
    intent: 'assist',
    prompt: 'Summarize this email or thread, keeping decisions and asks.',
    ground: ['this'],
    output: 'chat',
    contextMenu: true,
  },
  {
    id: 'catch-up',
    label: 'Catch me up',
    surfaces: ['outlook'],
    intent: 'assist',
    prompt: 'Catch me up on this conversation: what happened and what needs my attention.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'extract-actions',
    label: 'Extract action items & owners',
    surfaces: ['outlook'],
    intent: 'assist',
    prompt: 'Extract the action items and their owners from this email or thread.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'draft-reply',
    label: 'Draft a reply',
    surfaces: ['outlook'],
    intent: 'assist',
    prompt: 'Draft a reviewable reply to this email.',
    ground: ['this'],
    output: 'write',
    contextMenu: true,
  },
  {
    id: 'save-to-onenote',
    label: 'Summarize → save to OneNote',
    surfaces: ['outlook'],
    intent: 'synthesize',
    prompt: 'Summarize this email or thread and save the summary to a OneNote page.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },

  // ── Teams ────────────────────────────────────────────────────────────────────
  {
    id: 'live-notes',
    label: 'Live notes & recap',
    surfaces: ['teams'],
    intent: 'meeting-notes',
    prompt: 'Produce live notes and a recap from the meeting transcript.',
    ground: ['this'],
    output: 'annotation',
    contextMenu: false,
  },
  {
    id: 'action-items',
    label: 'Action items',
    surfaces: ['teams'],
    intent: 'meeting-notes',
    prompt: 'Extract action items and their owners from the meeting transcript.',
    ground: ['this'],
    output: 'annotation',
    contextMenu: false,
  },
  {
    id: 'catch-up-teams',
    label: 'Catch me up',
    surfaces: ['teams'],
    intent: 'assist',
    prompt: 'Catch me up on this meeting or channel: what happened and what needs my attention.',
    ground: ['this'],
    output: 'chat',
    contextMenu: false,
  },
  {
    id: 'post-summary',
    label: 'Post a summary to the channel',
    surfaces: ['teams'],
    intent: 'assist',
    prompt: 'Draft a reviewable summary to post to the channel.',
    ground: ['this'],
    output: 'write',
    contextMenu: false,
  },
];

/**
 * The quick actions offered on a surface, optionally narrowed by capability closure (ADR-0006).
 * Filters {@link QUICK_ACTIONS} to those listing `surface`, and — when `allowedIntents` is given —
 * drops any whose `intent` the surface cannot run, so the panel never offers an unreachable verb.
 */
export function quickActionsForSurface(
  surface: Surface,
  allowedIntents?: Iterable<Intent>,
): QuickAction[] {
  const allowed = allowedIntents ? new Set(allowedIntents) : undefined;
  return QUICK_ACTIONS.filter(
    (action) => action.surfaces.includes(surface) && (allowed ? allowed.has(action.intent) : true),
  );
}
