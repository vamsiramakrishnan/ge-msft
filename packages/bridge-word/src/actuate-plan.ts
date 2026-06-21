import type { ActuationRequest } from '@ge/contracts';

/**
 * Pure translation of an actuation into a host plan — testable without Office.js. Keeps the
 * content-anchoring rule explicit: a tracked change is located by `matchText` (+ contextHint)
 * via `body.search`, re-resolved at apply-time, never by a stored range id.
 */
export interface TrackedChangePlan {
  matchText?: string;
  contextHint?: string;
  contentControlId?: string;
  text: string;
}

export function planTrackedChange(req: ActuationRequest): TrackedChangePlan {
  const p = req.params;
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    ...(p.target?.contentControlId ? { contentControlId: p.target.contentControlId } : {}),
    text: p.text ?? '',
  };
}

/** Pick the search hit that matches the contextHint, else the first. Pure, so it's tested. */
export function chooseAnchorIndex(hitTexts: string[], contextHint?: string): number {
  if (contextHint) {
    const i = hitTexts.findIndex((t) => t.includes(contextHint));
    if (i >= 0) return i;
  }
  return hitTexts.length > 0 ? 0 : -1;
}
