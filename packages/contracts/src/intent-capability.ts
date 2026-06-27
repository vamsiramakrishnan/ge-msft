import type { Intent } from './intent.js';
import type { CapabilityManifest, ActuationKind } from './capability.js';

/**
 * Which actuation kinds an {@link Intent} needs to actually *land* its result on a surface. This is
 * the bridge between the general verbs (what the `/` palette and the quick actions dispatch) and the
 * capability closure (ADR-0006): a surface should only offer an intent whose required write it can
 * perform. The chat verbs (`ask`/`summarize`/`explain`) require nothing — they are grounded reads.
 *
 * An intent is satisfied if the manifest advertises **any** of its required kinds (e.g. a `rewrite`
 * can land as a tracked change OR a selection replace OR a cell write OR a filled content control).
 */
export const INTENT_REQUIRES: Record<Intent, ActuationKind[]> = {
  ask: [],
  summarize: [],
  explain: [],
  rewrite: ['tracked-change', 'replace-selection', 'write-cells', 'fill-content-control'],
  review: ['add-comment', 'tracked-change'],
  visualize: ['insert-chart'],
  draft: ['insert-slide', 'append-page', 'create-mail', 'reply-mail'],
  notes: ['post-message', 'post-card'],
};

/**
 * The intents a surface can honour, given its capability manifest: the chat verbs always, plus any
 * write/annotation verb at least one of whose required actuation kinds the manifest advertises. Pass
 * the result as `allowedIntents` to `quickActionsForSurface` / `commandPaletteFor` so the panel
 * never offers a verb the surface cannot run (the UI half of ADR-0006 closure — the runtime gate is
 * the other half).
 */
export function intentsForManifest(manifest: CapabilityManifest): Intent[] {
  const kinds = new Set<ActuationKind>(manifest.actuations.map((a) => a.kind));
  return (Object.keys(INTENT_REQUIRES) as Intent[]).filter((intent) => {
    const required = INTENT_REQUIRES[intent];
    return required.length === 0 || required.some((kind) => kinds.has(kind));
  });
}
