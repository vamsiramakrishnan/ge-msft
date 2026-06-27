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

/**
 * Pure plan for an `add-comment` actuation (ADR-0004 `comment` verb on Word): a NEW comment whose
 * body is `params.text`, anchored on `target.matchText` (+ `contextHint`) and re-resolved via
 * `body.search` at apply-time — the same content-anchoring discipline as a tracked change, so a
 * drifted anchor degrades rather than landing on the wrong range. The comment text is collapsed to
 * a single line (model/host-derived → untrusted) so it can't forge structure; `hasText` lets the
 * bridge reject an empty comment before touching the host.
 */
export interface AddCommentPlan {
  matchText?: string;
  contextHint?: string;
  text: string;
  hasText: boolean;
}

export function planAddComment(req: ActuationRequest): AddCommentPlan {
  const p = req.params;
  const text = oneLineSource(p.text ?? '');
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    text,
    hasText: text.length > 0,
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

/**
 * Pure plan for an `insert-text` actuation (ADR-0007 host-native write). Inserts plain `text` either
 * at a content anchor (`target.matchText` + `contextHint`, re-resolved via `body.search` at
 * apply-time — same discipline as a tracked change) or, when no anchor is given, at the current
 * selection. `anchored` distinguishes the two so the bridge can fail-closed on a missing anchor only
 * when one was *intended* but malformed, and otherwise fall through to the selection path. `hasText`
 * lets the bridge reject an empty insert before touching the host.
 */
export interface InsertTextPlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  text: string;
  hasText: boolean;
}

export function planInsertText(req: ActuationRequest): InsertTextPlan {
  const p = req.params;
  const text = p.text ?? '';
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    text,
    hasText: text.length > 0,
  };
}

/**
 * Pure plan for a `replace-selection` actuation (ADR-0007): overwrite the current selection with
 * `text`. There is no content anchor — the target is whatever the user has selected — so the plan is
 * just the text plus an `hasText` guard for the empty-insert fail-closed check.
 */
export interface ReplaceSelectionPlan {
  text: string;
  hasText: boolean;
}

export function planReplaceSelection(req: ActuationRequest): ReplaceSelectionPlan {
  const text = req.params.text ?? '';
  return { text, hasText: text.length > 0 };
}

/**
 * Pure plan for an `insert-ooxml` actuation (ADR-0007): insert rich `ooxml` either at a content
 * anchor (`target.matchText` + `contextHint`, re-resolved at apply-time) or at the current selection.
 * `anchored`/`hasOoxml` mirror {@link InsertTextPlan} so the bridge can apply the identical
 * fail-closed + degrade discipline. NOTE: the OOXML is host/model-derived → untrusted; the bridge
 * passes it to `range.insertOoxml` as data, never interpreting it.
 */
export interface InsertOoxmlPlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  ooxml: string;
  hasOoxml: boolean;
}

export function planInsertOoxml(req: ActuationRequest): InsertOoxmlPlan {
  const p = req.params;
  const ooxml = p.ooxml ?? '';
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    ooxml,
    hasOoxml: ooxml.length > 0,
  };
}

/**
 * Pure plan for a `fill-content-control` actuation (ADR-0007): populate the content control named by
 * `target.contentControlId` with `text`. Unlike the other kinds this is anchored by an explicit host
 * id (a content control is a stable, named container) rather than by matched content, so the plan
 * surfaces `contentControlId` + `hasId`. `hasText` guards the empty-fill case.
 */
export interface FillContentControlPlan {
  contentControlId?: string;
  hasId: boolean;
  text: string;
  hasText: boolean;
}

export function planFillContentControl(req: ActuationRequest): FillContentControlPlan {
  const p = req.params;
  const id = p.target?.contentControlId;
  const text = p.text ?? '';
  return {
    ...(id ? { contentControlId: id } : {}),
    hasId: Boolean(id),
    text,
    hasText: text.length > 0,
  };
}
