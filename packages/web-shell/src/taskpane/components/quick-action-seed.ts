import type { QuickAction } from '@ge/contracts';

/**
 * Render a {@link QuickAction} as the free-text seed handed to the assist session. The action's
 * default `@`-grounding sources (e.g. `['this']`, `['unit']`) are prepended as `@`-mentions ahead
 * of the templated `prompt`, so the seed reads exactly like something a user could have typed in
 * the composer (`@this Summarize this concisely…`). Pure and total — an action with no `ground`
 * degrades to just the prompt.
 */
export function quickActionSeed(action: QuickAction): string {
  const mentions = action.ground.map((g) => `@${g}`).join(' ');
  const prompt = action.prompt.trim();
  return mentions ? `${mentions} ${prompt}` : prompt;
}
