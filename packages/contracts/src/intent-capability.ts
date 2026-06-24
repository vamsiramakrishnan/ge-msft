import type { Intent } from './intent.js';
import type { CapabilityManifest, ActuationKind } from './capability.js';

/**
 * Which actuation kinds an {@link Intent} needs to actually *land* its result on a surface. This is
 * the bridge between the Intent verbs (what the `/` palette and the quick actions dispatch) and the
 * capability closure (ADR-0006): a surface should only offer an intent whose required write it can
 * perform. `assist` (and any read-only intent) requires nothing — it is grounded chat.
 *
 * An intent is satisfied if the manifest advertises **any** of its required kinds (e.g. `review`
 * findings can land as either a comment or a tracked change).
 */
export const INTENT_REQUIRES: Record<Intent, ActuationKind[]> = {
  assist: [],
  review: ['add-comment', 'tracked-change'],
  'resolve-comment': ['comment-reply'],
  'regen-clause': ['tracked-change', 'fill-content-control', 'replace-selection'],
  'draft-slides': ['insert-slide'],
  synthesize: ['append-page'],
  'meeting-notes': ['post-message', 'post-card'],
};

/**
 * The intents a surface can honour, given its capability manifest: `assist` always, plus any intent
 * at least one of whose required actuation kinds the manifest advertises. Pass the result as
 * `allowedIntents` to `quickActionsForSurface` / `commandPaletteFor` so the panel never offers a
 * verb the surface cannot run (the UI half of ADR-0006 closure — the runtime gate is the other half).
 */
export function intentsForManifest(manifest: CapabilityManifest): Intent[] {
  const kinds = new Set<ActuationKind>(manifest.actuations.map((a) => a.kind));
  return (Object.keys(INTENT_REQUIRES) as Intent[]).filter((intent) => {
    const required = INTENT_REQUIRES[intent];
    return required.length === 0 || required.some((kind) => kinds.has(kind));
  });
}
