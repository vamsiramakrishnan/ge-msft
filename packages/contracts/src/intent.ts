import { z } from 'zod';

/**
 * The general, Copilot-altitude verbs the assistant handles (EXPERIENCE.md §1). The verb is the
 * WHAT — kept small and surface-agnostic; the WHERE is the orthogonal {@link CommandScope}, the
 * grounding is the orthogonal {@link GroundSource}. Chat verbs (`ask`/`summarize`/`explain`) are
 * single-shot reads routed to `send`; the four specialist verbs (`rewrite`/`review`/`draft`/
 * `notes`) fan out to a confirmable plan and route through `runCommands` (the gate).
 *
 * - `ask`       — grounded chat / a custom free-text prompt over a scope (the rename of `assist`).
 * - `summarize` — condense the scope.
 * - `explain`   — clarify the scope in plain language.
 * - `rewrite`   — apply any instruction to the scope → a reversible edit.
 * - `review`    — whole-scope pass emitting N findings → N gated annotations.
 * - `draft`     — generate new host material (slides, a page, a reply, a column).
 * - `notes`     — transcript → live notes + action items (Teams).
 */
export const IntentSchema = z.enum([
  'ask', // grounded chat / custom prompt over a scope (StreamAssist)
  'summarize', // condense the scope
  'explain', // clarify the scope in plain language
  'rewrite', // apply any instruction to the scope -> a reversible edit
  'review', // whole-scope pass -> N findings -> N gated annotations
  'draft', // generate new host material (slides / page / reply / column)
  'notes', // transcript -> live notes + action items (Teams)
]);

export type Intent = z.infer<typeof IntentSchema>;

/**
 * SCOPE (Tier 2, EXPERIENCE.md §1) — *where* a verb acts, a first-class orthogonal axis, never a
 * verb. Surfaced as a segmented control next to Send; the surface-named labels are supplied as data
 * in the palette (`CommandPaletteSpec.scopeOptions`) so the core stays surface-agnostic. `ref`
 * carries the addressable target for `range` (A1/named), `section` (heading), and `comment` (id).
 */
export const CommandScopeSchema = z.object({
  kind: z.enum(['selection', 'document', 'range', 'section', 'comment', 'this-item']),
  ref: z.string().optional(),
});
export type CommandScope = z.infer<typeof CommandScopeSchema>;

/**
 * GROUND (Tier 3, EXPERIENCE.md §1) — *what a turn is grounded on*, the `@`-mention picker output.
 * `ground` and the old `mentionKinds` are the same concept, unified here into one typed vocabulary:
 *
 * - `this`      — the live scope ("this {selection|range|slide|thread}").
 * - `unit`      — the research unit (notebook + federated sources + working document).
 * - `document`  — a named document already in a connected data store.
 * - `person`    — a person/contact reference.
 * - `datastore` — a whole connected data store.
 * - `upload`    — an ad-hoc uploaded source.
 */
export const GroundSourceSchema = z.enum([
  'this',
  'unit',
  'document',
  'person',
  'datastore',
  'upload',
]);
export type GroundSource = z.infer<typeof GroundSourceSchema>;
