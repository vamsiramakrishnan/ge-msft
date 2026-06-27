import { fillPrompt, type CommandScope, type QuickAction } from '@ge/contracts';
import type { ComposerInvocation, ComposerMention } from './Composer.js';

/**
 * The deterministic STRING seed a typed invocation compiles to for the controller's `send` /
 * `runCommands` (both still take a `string` task — the cmd executor + plan grammar are unchanged).
 *
 * We no longer stringify a {@link QuickAction} into an `@`-magic-string at the call site
 * (EXPERIENCE.md §3): a chip / a `/verb` flows as a typed {@link ComposerInvocation}, routed by
 * `output` (chat → `send`, write/annotation → `runCommands`). Only HERE, at the controller seam, is
 * the typed tuple rendered to a string — deterministically from the typed fields (mentions ahead of
 * the instruction, exactly like a line a user could have typed), so a planted/free-text seed can
 * never widen grounding.
 */

/** Render the typed mentions as `@kind`/`@kind:ref` tokens, in order (deterministic). */
function renderMentions(mentions: readonly ComposerMention[]): string {
  return mentions.map((m) => (m.ref ? `@${m.kind}:${m.ref}` : `@${m.kind}`)).join(' ');
}

/**
 * Compile a {@link ComposerInvocation} to the controller's string seed.
 *
 * - A composer-origin invocation carries the verbatim `raw` (the user already typed `/verb`, inline
 *   `@mentions`, and the instruction as ONE well-formed line) — we use it as the body so we never
 *   double-render a mention the user already typed inline.
 * - A chip-origin invocation has an empty `raw` (its mentions live in the typed `mentions` field, not
 *   inline) — we synthesize the body deterministically as `/verb @mentions instruction`.
 *
 * Either way, scope rides as a typed field (from the segmented control, never typed into the box);
 * when it is not the live `selection`, a `scope:` token is appended so the plan grammar — which
 * already threads `scope` — picks it up. The composer never types `scope:`, so the append is never a
 * duplicate.
 */
export function invocationToSeed(inv: ComposerInvocation): string {
  const body =
    inv.raw.trim().length > 0
      ? inv.raw.trim()
      : [inv.intent ? `/${inv.intent}` : '', renderMentions(inv.mentions), inv.instruction.trim()]
          .filter((p) => p.length > 0)
          .join(' ');
  return [body, scopeToken(inv.scope)].filter((p) => p.length > 0).join(' ');
}

/** Render a scope as a `scope:` token, omitting the live `selection` default (it is implicit). */
function scopeToken(scope: CommandScope): string {
  if (scope.kind === 'selection') return '';
  return scope.ref ? `scope:${scope.kind}(${scope.ref})` : `scope:${scope.kind}`;
}

/**
 * Turn a {@link QuickAction} into a typed {@link ComposerInvocation} — the SAME typed shape a
 * composer submit produces, so a chip and a typed line share one downstream path. The action's
 * `ground` defaults become typed mentions and its `scope`/`intent` carry through verbatim. A
 * parameterized action's `{{name}}` slots are substituted from `values` (Workstream H) BEFORE the
 * instruction is built — the panel collects every value first, so a raw `{{topic}}` never reaches the
 * model; an unprovided slot is left intact for the fail-closed dispatch guard to catch.
 */
export function quickActionToInvocation(
  action: QuickAction,
  values: Readonly<Record<string, string>> = {},
): ComposerInvocation {
  const mentions: ComposerMention[] = action.ground.map((g) => ({
    kind: g as ComposerMention['kind'],
  }));
  // `raw` is empty: a chip is a typed-field origin (its mentions live in `mentions`, not inline), so
  // `invocationToSeed` synthesizes the line deterministically rather than echoing a verbatim string.
  return {
    intent: action.intent,
    scope: { ...action.scope },
    mentions,
    instruction: fillPrompt(action.prompt, values).trim(),
    raw: '',
  };
}

/**
 * The controller string seed for a {@link QuickAction} — built deterministically from its typed
 * fields via {@link quickActionToInvocation} → {@link invocationToSeed}. Pure and total.
 */
export function quickActionSeed(action: QuickAction): string {
  return invocationToSeed(quickActionToInvocation(action));
}
