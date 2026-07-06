import { z } from 'zod';
import { IntentSchema, CommandScopeSchema, type Intent, type CommandScope } from './intent.js';
import { SurfaceSchema, type Surface } from './context.js';
import { INTENT_REQUIRES } from './intent-capability.js';
import { VERBS_BY_SURFACE } from './command-palette.js';

/**
 * Quick actions — the curated, one-tap verb catalog the panel and the right-click context menu
 * offer per surface. Each entry is a legible **preset tuple** (EXPERIENCE.md §5): a general `intent`,
 * a `scope`, a default `@`-grounding set, and a templated free-text `prompt` — the same artifact the
 * planner emits and the executor runs. It is scoped by capability closure (ADR-0006); an action whose
 * intent the surface cannot run is never offered.
 *
 * `output` is **derivable, not declared**: a write/annotation verb is exactly one with a non-empty
 * {@link INTENT_REQUIRES} entry (the rule that closes the silent drift the audit found in
 * `draft-reply`/`risk-column`/`write-formula`). Every action's `output` is set by {@link deriveOutput}
 * from its intent, so the catalog can never drift from the closure.
 *
 * `ground` defaults: read-only chat actions ground on `['this']` (the live scope); actions that lean
 * on the research unit ground on `['unit']` (the notebook + federated sources + working document).
 */

/** The output shape an action produces: grounded chat, an inline annotation pass, or a write-back. */
export const QuickActionOutputSchema = z.enum(['chat', 'annotation', 'write']);
export type QuickActionOutput = z.infer<typeof QuickActionOutputSchema>;

/**
 * A fill-in slot in an action's `prompt`, referenced as `{{name}}` (EXPERIENCE.md §5, Workstream H).
 * A parameterized action is NEVER dispatched with its template literal — the panel collects every
 * declared value first, so a raw `{{topic}}` can never reach the model. `name` is the token; `label`
 * is what the fill field shows; `hint` is an example placeholder. `name` must be a clean token (no
 * braces/whitespace) so the parity check below can match it against the template.
 */
export const QuickActionParamSchema = z.object({
  name: z
    .string()
    .regex(/^[a-z][a-z0-9_]*$/i, 'parameter name must be a bare token (no braces or spaces)'),
  label: z.string().min(1),
  hint: z.string().optional(),
});
export type QuickActionParam = z.infer<typeof QuickActionParamSchema>;

/** Every distinct `{{name}}` token in a prompt, in order of first appearance (deduped). */
export function promptPlaceholders(prompt: string): string[] {
  const seen = new Set<string>();
  for (const m of prompt.matchAll(/\{\{\s*([^}]+?)\s*\}\}/g)) {
    const name = (m[1] ?? '').trim();
    if (name) seen.add(name);
  }
  return [...seen];
}

/** The declared parameters of an action, normalized to an array (the field is optional). */
export function actionParameters(action: Pick<QuickAction, 'parameters'>): QuickActionParam[] {
  return action.parameters ?? [];
}

/** Substitute every `{{name}}` in `prompt` with `values[name]`; unprovided names are left intact. */
export function fillPrompt(prompt: string, values: Readonly<Record<string, string>>): string {
  return prompt.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (whole, raw: string) => {
    const name = raw.trim();
    return Object.prototype.hasOwnProperty.call(values, name) ? values[name]! : whole;
  });
}

/** True iff `text` still carries an unfilled `{{…}}` placeholder — the fail-closed dispatch guard. */
export function hasUnfilledPlaceholder(text: string): boolean {
  return /\{\{\s*[^}]+?\s*\}\}/.test(text);
}

export const QuickActionSchema = z
  .object({
    id: z.string(),
    label: z.string(),
    surfaces: z.array(SurfaceSchema).nonempty(),
    intent: IntentSchema,
    scope: CommandScopeSchema, // WHERE the action acts (EXPERIENCE.md §1, Tier 2)
    prompt: z.string(), // templated free-text seed; `{{name}}` slots are declared in `parameters`
    ground: z.array(z.string()).default([]), // default @-sources, e.g. ['this']
    output: QuickActionOutputSchema, // MUST equal deriveOutput(intent) — enforced below
    contextMenu: z.boolean().default(false), // also offered on right-click
    parameters: z.array(QuickActionParamSchema).optional(), // typed `{{name}}` fill slots (H)
  })
  // Fail-closed: an action's declared `output` MUST match the closure-derived output for its intent.
  // The panel routes the gate off the intent (write/annotation → gate), so a tenant-composed catalog
  // entry with a mismatched output (e.g. {intent:'ask', output:'write'}) is a safety trap; reject it
  // at parse time rather than only in the factory helpers. (Security review, Finding 1.)
  .superRefine((a, ctx) => {
    const expected = deriveOutput(a.intent);
    if (a.output !== expected) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ['output'],
        message: `output '${a.output}' must equal deriveOutput('${a.intent}') = '${expected}'`,
      });
    }
    // Template ↔ parameter parity (H): every `{{name}}` slot in the prompt MUST be a declared
    // parameter, and every declared parameter MUST appear in the prompt. This makes it structurally
    // impossible to ship an action that leaks a literal `{{…}}` to the model (an undeclared slot) or
    // that prompts for a value it never uses (a dangling param).
    const slots = new Set(promptPlaceholders(a.prompt));
    const params = new Set((a.parameters ?? []).map((p) => p.name));
    for (const slot of slots) {
      if (!params.has(slot)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['parameters'],
          message: `prompt slot '{{${slot}}}' has no declared parameter`,
        });
      }
    }
    for (const name of params) {
      if (!slots.has(name)) {
        ctx.addIssue({
          code: z.ZodIssueCode.custom,
          path: ['parameters'],
          message: `parameter '${name}' is not referenced as '{{${name}}}' in the prompt`,
        });
      }
    }
  });
export type QuickAction = z.infer<typeof QuickActionSchema>;

/** All six surfaces — generated from the schema so adding a surface needs no edit here. */
const ALL_SURFACES = SurfaceSchema.options as [Surface, ...Surface[]];

/**
 * Derive the `output` shape from the intent: a verb with no required actuation is a chat read; a
 * `review`/`notes` verb is an annotation pass; any other write verb (`rewrite`/`draft`) lands a
 * write-back. This is the single source of truth the catalog's literal `output` is asserted against.
 */
export function deriveOutput(intent: Intent): QuickActionOutput {
  if (INTENT_REQUIRES[intent].length === 0) return 'chat';
  return intent === 'review' || intent === 'notes' ? 'annotation' : 'write';
}

/** Two common scopes, spelled once. */
const THIS_ITEM: CommandScope = { kind: 'this-item' };
const SELECTION: CommandScope = { kind: 'selection' };
const DOCUMENT: CommandScope = { kind: 'document' };

/** A universal action seed (its `surfaces` and `output` are filled in by {@link universal}). */
type UniversalSeed = Omit<QuickAction, 'surfaces' | 'output'>;

/**
 * Expand a universal seed across every `Surface` (EXPERIENCE.md §4), keeping only the surfaces whose
 * palette actually offers the seed's intent (so a universal `explain` lands on every surface that
 * has `/explain`, but not on a surface like Teams whose verb set omits it). This keeps the catalog
 * within capability closure by construction — no surface ever lists an action it cannot run.
 */
function universal(seed: UniversalSeed): QuickAction {
  const surfaces = ALL_SURFACES.filter((s) => VERBS_BY_SURFACE[s].includes(seed.intent));
  if (surfaces.length === 0) {
    throw new Error(
      `universal action '${seed.id}' uses intent '${seed.intent}' offered on no surface`,
    );
  }
  return {
    ...seed,
    surfaces: surfaces as [Surface, ...Surface[]],
    output: deriveOutput(seed.intent),
  };
}

/** A surface action seed (its `output` is derived by {@link surfaceAction}). */
type SurfaceSeed = Omit<QuickAction, 'output'>;

/** Build a surface action, deriving its `output` from the intent. */
function surfaceAction(seed: SurfaceSeed): QuickAction {
  return { ...seed, output: deriveOutput(seed.intent) };
}

/**
 * The universal block — generated for **every** `Surface` from `SurfaceSchema.options`
 * (EXPERIENCE.md §4). Adding a surface needs no edit here. All four are chat reads grounded on the
 * live scope.
 */
export const UNIVERSAL_ACTIONS: QuickAction[] = [
  universal({
    id: 'summarize-this',
    label: 'Summarize this',
    intent: 'summarize',
    scope: THIS_ITEM,
    prompt: 'Summarize this concisely, keeping the key facts and figures.',
    ground: ['this'],
    contextMenu: true,
  }),
  universal({
    id: 'key-points',
    label: 'Key points & action items',
    intent: 'ask',
    scope: THIS_ITEM,
    prompt: 'List the key points and any action items, with owners where stated.',
    ground: ['this'],
    contextMenu: false,
  }),
  universal({
    id: 'explain',
    label: 'Explain / clarify',
    intent: 'explain',
    scope: THIS_ITEM,
    prompt: 'Explain this in plain language and clarify anything ambiguous.',
    ground: ['this'],
    contextMenu: true,
  }),
  universal({
    id: 'find-risks',
    label: 'Find risks & gaps',
    intent: 'ask',
    scope: THIS_ITEM,
    prompt: 'Identify the risks, gaps, and open questions in this.',
    ground: ['this'],
    contextMenu: false,
  }),
];

/** The genuine per-surface specializations (EXPERIENCE.md §2). */
export const SURFACE_ACTIONS: QuickAction[] = [
  // ── Word ───────────────────────────────────────────────────────────────────
  surfaceAction({
    id: 'review-against',
    label: 'Review against…',
    surfaces: ['word'],
    intent: 'review',
    scope: DOCUMENT,
    // A general review against a named standard (policy is just one possible target — the
    // contract-review nouns moved to CONTRACT_REVIEW_PACK).
    prompt: 'Review this document against {{standard}} and flag every breach as a finding.',
    ground: ['unit'],
    contextMenu: false,
    parameters: [
      { name: 'standard', label: 'Review against', hint: 'e.g. the master agreement, ISO 27001' },
    ],
  }),
  surfaceAction({
    id: 'tighten',
    label: 'Tighten / rewrite selection',
    surfaces: ['word'],
    intent: 'rewrite',
    scope: SELECTION,
    prompt: 'Tighten and rewrite the selected text as a tracked change, preserving its meaning.',
    ground: ['this'],
    contextMenu: true,
  }),
  surfaceAction({
    id: 'exec-summary',
    label: 'Draft an executive summary',
    surfaces: ['word'],
    intent: 'summarize',
    scope: DOCUMENT,
    prompt: 'Draft a concise executive summary of this document.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'resolve-comment',
    label: 'Resolve this comment',
    surfaces: ['word'],
    intent: 'rewrite',
    scope: { kind: 'comment' },
    prompt: 'Resolve this comment: edit the anchored text as a tracked change.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'comment-on-issues',
    label: 'Comment on issues',
    surfaces: ['word'],
    intent: 'review',
    scope: DOCUMENT,
    prompt: 'Review this document and add comments for the most important issues.',
    ground: ['this'],
    contextMenu: false,
  }),

  // ── Excel ────────────────────────────────────────────────────────────────────
  surfaceAction({
    // READ-ONLY BY CONTRACT: `intent: 'ask'` → `output: 'chat'` → routes to `send`, never a cell
    // write. So the `=GE.ASK("…")` text is only the model-facing task, and `fillPrompt`'s naked
    // (unescaped) substitution into the formula string is safe. If a formula-shaped prompt ever gains
    // a WRITE path, add a formula-slot encoder (escape `"`, reject newlines) first (security review H).
    id: 'ge-ask',
    label: '=GE.ASK in the grid',
    surfaces: ['excel'],
    intent: 'ask',
    scope: { kind: 'range' },
    prompt: '=GE.ASK("{{question}}", {{range}})',
    ground: ['this'],
    contextMenu: false,
    parameters: [
      { name: 'question', label: 'Question', hint: 'e.g. which region grew fastest?' },
      { name: 'range', label: 'Cell range', hint: 'e.g. A1:D20' },
    ],
  }),
  surfaceAction({
    id: 'summarize-range',
    label: 'Summarize this range',
    surfaces: ['excel'],
    intent: 'summarize',
    scope: SELECTION,
    prompt: 'Summarize the selected range, calling out totals, trends, and outliers.',
    ground: ['this'],
    contextMenu: true,
  }),
  surfaceAction({
    id: 'explain-formula',
    label: 'Explain this formula',
    surfaces: ['excel'],
    intent: 'explain',
    scope: SELECTION,
    prompt: 'Explain what the formula in the selected cell does, step by step.',
    ground: ['this'],
    contextMenu: true,
  }),
  surfaceAction({
    id: 'risk-column',
    label: 'Add a risk/summary column',
    surfaces: ['excel'],
    intent: 'rewrite',
    scope: { kind: 'range' },
    prompt: 'Add a risk or summary column derived from the selected range.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'find-anomalies',
    label: 'Find anomalies / outliers',
    surfaces: ['excel'],
    intent: 'review',
    scope: SELECTION,
    prompt: 'Find anomalies and outliers in the selected range and comment on each.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'create-chart',
    label: 'Create a chart',
    surfaces: ['excel'],
    intent: 'visualize',
    scope: { kind: 'range' },
    prompt:
      'Create an appropriate chart from the selected range. Read the range first, choose a supported chart type, and use the chart command.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'write-formula',
    label: 'Write a formula for…',
    surfaces: ['excel'],
    intent: 'rewrite',
    scope: { kind: 'range' },
    prompt: 'Write a formula for {{goal}} and place it in the target cell.',
    ground: ['this'],
    contextMenu: false,
    parameters: [
      { name: 'goal', label: 'Compute', hint: 'e.g. variance of column B against forecast' },
    ],
  }),

  // ── PowerPoint ───────────────────────────────────────────────────────────────
  surfaceAction({
    id: 'draft-section',
    label: 'Draft a section from the unit',
    surfaces: ['powerpoint'],
    intent: 'draft',
    scope: DOCUMENT,
    prompt: 'Draft a slide section from the research unit on {{topic}}.',
    ground: ['unit'],
    contextMenu: false,
    parameters: [{ name: 'topic', label: 'Topic', hint: 'e.g. Q3 go-to-market strategy' }],
  }),
  surfaceAction({
    id: 'draft-slide',
    label: 'Draft a slide',
    surfaces: ['powerpoint'],
    intent: 'draft',
    scope: THIS_ITEM,
    prompt: 'Draft one new slide on {{topic}}, with a clear title and concise bullets.',
    ground: ['unit'],
    contextMenu: false,
    parameters: [{ name: 'topic', label: 'Slide topic', hint: 'e.g. Q4 outlook' }],
  }),
  surfaceAction({
    id: 'speaker-notes',
    label: 'Draft speaker notes',
    surfaces: ['powerpoint'],
    // No host write path for speaker notes yet (CAPABILITY-MAP: set-speaker-notes is
    // modeled-not-advertised) — draft them into the pane as chat until the bridge can actuate.
    intent: 'ask',
    scope: THIS_ITEM,
    prompt: 'Draft speaker notes for the current slide.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'summarize-deck',
    label: 'Summarize the deck',
    surfaces: ['powerpoint'],
    intent: 'summarize',
    scope: DOCUMENT,
    prompt: 'Summarize the narrative and key takeaways of this deck.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'redesign',
    label: 'Suggest a redesign',
    surfaces: ['powerpoint'],
    intent: 'ask',
    scope: THIS_ITEM,
    prompt: 'Suggest a redesign for the current slide to make it clearer and tighter.',
    ground: ['this'],
    contextMenu: false,
  }),

  // ── OneNote ──────────────────────────────────────────────────────────────────
  surfaceAction({
    id: 'synthesize-page',
    label: 'Summarize sources onto the page',
    surfaces: ['onenote'],
    intent: 'draft',
    scope: DOCUMENT,
    prompt: 'Synthesize the grounded sources into a cited summary on this page.',
    ground: ['unit'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'add-sources-to-unit',
    label: "Add this page's sources to the unit",
    surfaces: ['onenote'],
    // OneNote is where the unit is assembled (EXPERIENCE.md §2) — a staged composition action;
    // it adds to the unit rather than writing host material, so it stays a chat/staged action.
    intent: 'ask',
    scope: DOCUMENT,
    prompt: "Add this page's linked sources to the research unit.",
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'audio-overview',
    // Honest label: this produces a SCRIPT in chat (output:'chat'), not rendered audio — the copy
    // must not promise audio the surface cannot deliver (EXPERIENCE.md §2: stays `output:'chat'`).
    label: 'Draft an audio-overview script',
    surfaces: ['onenote'],
    intent: 'ask',
    scope: DOCUMENT,
    prompt: 'Draft an audio-overview script of the grounded sources.',
    ground: ['unit'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'discover-sources',
    label: 'Discover related sources',
    surfaces: ['onenote'],
    intent: 'ask',
    scope: DOCUMENT,
    prompt: 'Discover sources related to this page from the connected data stores.',
    ground: ['this'],
    contextMenu: false,
  }),

  // ── Outlook ──────────────────────────────────────────────────────────────────
  surfaceAction({
    id: 'summarize-email',
    label: 'Summarize this email / thread',
    surfaces: ['outlook'],
    intent: 'summarize',
    scope: THIS_ITEM,
    prompt: 'Summarize this email or thread, keeping decisions and asks.',
    ground: ['this'],
    contextMenu: true,
  }),
  surfaceAction({
    id: 'catch-up',
    label: 'Catch me up',
    surfaces: ['outlook'],
    intent: 'summarize',
    scope: DOCUMENT,
    prompt: 'Catch me up on this conversation: what happened and what needs my attention.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'extract-actions',
    label: 'Extract action items & owners',
    surfaces: ['outlook'],
    intent: 'ask',
    scope: DOCUMENT,
    prompt: 'Extract the action items and their owners from this email or thread.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'draft-reply',
    label: 'Draft a reply',
    surfaces: ['outlook'],
    intent: 'draft',
    scope: DOCUMENT,
    prompt: 'Draft a reviewable reply to this thread.',
    ground: ['this'],
    contextMenu: true,
  }),
  surfaceAction({
    id: 'draft-reply-toned',
    label: 'Draft a reply in a tone…',
    surfaces: ['outlook'],
    // Tone control on reply drafting (EXPERIENCE.md §2 — Copilot has tone; we expose it as a
    // typed {{tone}} fill slot). The one-tap `draft-reply` above stays untouched.
    intent: 'draft',
    scope: DOCUMENT,
    prompt:
      'Draft a reply to this thread in a {{tone}} tone; keep commitments accurate and questions answered.',
    ground: ['this'],
    contextMenu: false,
    parameters: [{ name: 'tone', label: 'Tone', hint: 'e.g. formal, warm, brief' }],
  }),
  surfaceAction({
    id: 'draft-new-email',
    label: 'Draft a new email',
    surfaces: ['outlook'],
    intent: 'draft',
    scope: DOCUMENT,
    prompt:
      'Draft a new email with subject {{subject}}. Leave recipients empty unless I provide them.',
    ground: ['unit'],
    contextMenu: false,
    parameters: [{ name: 'subject', label: 'Subject', hint: 'e.g. Follow-up on Q3 planning' }],
  }),

  // ── Teams ────────────────────────────────────────────────────────────────────
  surfaceAction({
    id: 'live-notes',
    label: 'Live notes & recap',
    surfaces: ['teams'],
    intent: 'notes',
    scope: DOCUMENT,
    prompt: 'Produce live notes and a recap from the meeting transcript.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'action-items',
    label: 'Action items',
    surfaces: ['teams'],
    intent: 'notes',
    scope: DOCUMENT,
    prompt: 'Extract action items and their owners from the meeting transcript.',
    ground: ['this'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'catch-up-teams',
    label: 'Catch me up',
    surfaces: ['teams'],
    intent: 'summarize',
    scope: DOCUMENT,
    prompt: 'Catch me up on this meeting or channel: what happened and what needs my attention.',
    ground: ['this'],
    contextMenu: false,
  }),
];

/**
 * The quick-action catalog: the universal block followed by the per-surface specializations. Every
 * entry's `output` is derived from its intent (no hand-declared drift).
 */
export const QUICK_ACTIONS: QuickAction[] = [...UNIVERSAL_ACTIONS, ...SURFACE_ACTIONS];

/**
 * An optional vertical pack: the contract-review nouns that used to sit in the default Word catalog
 * (`review-policy` / `find-unsupported`). Kept out of {@link QUICK_ACTIONS} so the default catalog
 * stays general; a tenant that wants them composes `[...QUICK_ACTIONS, ...CONTRACT_REVIEW_PACK]`.
 */
export const CONTRACT_REVIEW_PACK: QuickAction[] = [
  surfaceAction({
    id: 'review-policy',
    label: 'Review against policy',
    surfaces: ['word'],
    intent: 'review',
    scope: DOCUMENT,
    prompt: 'Review this document against the grounded policy and flag every breach as a finding.',
    ground: ['unit'],
    contextMenu: false,
  }),
  surfaceAction({
    id: 'find-unsupported',
    label: 'Find unsupported claims',
    surfaces: ['word'],
    intent: 'review',
    scope: DOCUMENT,
    prompt: 'Flag claims that are unsupported by the grounded sources.',
    ground: ['unit'],
    contextMenu: false,
  }),
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
