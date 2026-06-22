import type { ActuationRequest, SourceRef } from '@ge/contracts';

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

/** Collapse control chars + whitespace to one line so a source can't forge extra comment lines. */
function oneLineSource(text: string): string {
  const stripped = Array.from(text, (ch) => {
    const cp = ch.codePointAt(0) ?? 0;
    return cp < 0x20 || cp === 0x7f ? ' ' : ch;
  }).join('');
  return stripped.replace(/\s+/g, ' ').trim();
}

/** Render one source as `Title (uri)`, dropping any non-http(s) uri (no `javascript:`/`data:`). */
function renderSource(s: SourceRef): string {
  const title = oneLineSource(s.title);
  const uri = s.uri && /^https?:\/\//i.test(s.uri.trim()) ? oneLineSource(s.uri) : undefined;
  return uri ? `${title} (${uri})` : title;
}

/**
 * Render a source list as the text of a citation comment (ADR-0003 comments-as-citations). One
 * `Title (uri)` per line — the uri is dropped when absent or not http(s), each source single-lined
 * so a crafted title can't forge extra lines, and the whole thing capped at `maxChars`. Empty list
 * → empty string, which the bridge treats as "no comment to add". Pure, so it's unit-tested.
 */
export function formatSources(sources: SourceRef[], maxChars = 1000): string {
  const text = sources.map(renderSource).join('\n');
  if (text.length <= maxChars) return text;
  if (maxChars <= 1) return text.slice(0, maxChars);
  return `${text.slice(0, maxChars - 1)}…`;
}

/** Pick the search hit that matches the contextHint, else the first. Pure, so it's tested. */
export function chooseAnchorIndex(hitTexts: string[], contextHint?: string): number {
  if (contextHint) {
    const i = hitTexts.findIndex((t) => t.includes(contextHint));
    if (i >= 0) return i;
  }
  return hitTexts.length > 0 ? 0 : -1;
}
