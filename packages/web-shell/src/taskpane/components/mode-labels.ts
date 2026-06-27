import type { QuickAction } from '@ge/contracts';

/** The three output kinds an action can produce: a read-only answer, a reviewable annotation, or a write. */
export type OutputKind = QuickAction['output'];

/**
 * ONE canonical lexicon for the three output modes, consumed everywhere a mode is labelled
 * (surface command center, quick-action drawer, parameter form). Before this, each component
 * invented its own names — Ask/Review/Change vs Answer/Review gate/Preview gate vs
 * Ask/Preview comments/Preview write — so the same mode read three different ways. Keep all
 * mode wording here so the surface speaks one dialect of the grammar the model emits.
 */

/** Short noun-ish mode name for chips, tabs, and the command-center badge. */
export const MODE_LABEL: Record<OutputKind, string> = {
  chat: 'Ask',
  annotation: 'Review',
  write: 'Change',
};

/** Imperative button/CTA wording — what acting on the mode will do, gate-aware. */
export const MODE_CTA: Record<OutputKind, string> = {
  chat: 'Ask',
  annotation: 'Preview comments',
  write: 'Preview change',
};

export function modeLabel(action: Pick<QuickAction, 'output'>): string {
  return MODE_LABEL[action.output];
}

export function modeCta(action: Pick<QuickAction, 'output'>): string {
  return MODE_CTA[action.output];
}
