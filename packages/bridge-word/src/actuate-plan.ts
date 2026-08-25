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

/**
 * Pure plan for an `apply-style` actuation: set a named style on the resolved range. A built-in
 * style writes `styleBuiltIn` (portable across locales, WordApi 1.3); otherwise the localized
 * `style` name is written (WordApi 1.1). Anchoring mirrors {@link InsertTextPlan}.
 */
export interface ApplyStylePlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  styleName: string;
  builtIn: boolean;
  hasStyle: boolean;
}

export function planApplyStyle(req: ActuationRequest): ApplyStylePlan {
  const p = req.params;
  const styleName = p.style?.name ?? '';
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    styleName,
    builtIn: p.style?.builtIn ?? false,
    hasStyle: styleName.length > 0,
  };
}

/**
 * Pure plan for an `insert-table` actuation: build a native table from the `tableGrid` value grid
 * at the selection or after an anchor. `hasRows` fails closed on an empty/ragged-empty grid before
 * touching the host; the row/column counts derive from the grid itself.
 */
export interface InsertTablePlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  rows: string[][];
  hasRows: boolean;
}

export function planInsertTable(req: ActuationRequest): InsertTablePlan {
  const p = req.params;
  const rows = p.tableGrid?.rows ?? [];
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    rows,
    hasRows: rows.length > 0 && (rows[0]?.length ?? 0) > 0,
  };
}

/**
 * Pure plan for an `insert-content-control` actuation: wrap the selection or anchored range in a NEW
 * content control with an optional type/tag/title. The descriptor itself is required (the bridge
 * fails closed without it); every field inside it is host-applied verbatim.
 */
export interface InsertContentControlPlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  controlType?: string;
  tag?: string;
  title?: string;
  hasControl: boolean;
}

export function planInsertContentControl(req: ActuationRequest): InsertContentControlPlan {
  const p = req.params;
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    ...(p.contentControl?.type ? { controlType: p.contentControl.type } : {}),
    ...(p.contentControl?.tag ? { tag: p.contentControl.tag } : {}),
    ...(p.contentControl?.title ? { title: p.contentControl.title } : {}),
    hasControl: p.contentControl !== undefined,
  };
}

/**
 * Pure plan for an `insert-hyperlink` actuation: point the resolved range's hyperlink at `url`.
 * NOTE: the URL is model/host-derived → untrusted; the bridge screens it against http(s) (the same
 * allowlist discipline as {@link formatSources}) before it ever reaches the host.
 */
export interface InsertHyperlinkPlan {
  matchText?: string;
  contextHint?: string;
  anchored: boolean;
  url: string;
  hasUrl: boolean;
}

export function planInsertHyperlink(req: ActuationRequest): InsertHyperlinkPlan {
  const p = req.params;
  const url = p.hyperlink?.url ?? '';
  return {
    ...(p.target?.matchText ? { matchText: p.target.matchText } : {}),
    ...(p.target?.contextHint ? { contextHint: p.target.contextHint } : {}),
    anchored: Boolean(p.target?.matchText),
    url,
    hasUrl: url.length > 0,
  };
}

/**
 * Pure plan for a `find-replace` actuation: replace every occurrence of exact text across the body.
 * `replace` MAY be empty (deleting all hits); only an absent descriptor or empty `find` fails
 * closed. Case/whole-word options pass through to `body.search` verbatim.
 */
export interface FindReplacePlan {
  find: string;
  replace: string;
  matchCase: boolean;
  matchWholeWord: boolean;
  hasFindReplace: boolean;
}

export function planFindReplace(req: ActuationRequest): FindReplacePlan {
  const p = req.params;
  return {
    find: p.findReplace?.find ?? '',
    replace: p.findReplace?.replace ?? '',
    matchCase: p.findReplace?.matchCase ?? false,
    matchWholeWord: p.findReplace?.matchWholeWord ?? false,
    hasFindReplace: p.findReplace !== undefined && (p.findReplace.find ?? '').trim().length > 0,
  };
}
