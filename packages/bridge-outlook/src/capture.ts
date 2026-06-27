import type { Block } from '@ge/content';
import type { ResolvedContext } from '@ge/contracts';
import { native, toContext, type ToContextOptions } from '@ge/content';

/**
 * Pure mapping from an Outlook mail item into grounding-ready context — no Office.js here,
 * so it's unit-testable. Unlike Word/Excel, Outlook has no native block model: the body
 * arrives as an HTML/text **string**, so we take @ge/content's string path. The `OutlookBridge`
 * reads subject/from/body via the callback-based `getAsync` APIs and hands the plain values
 * to this function, which labels them (subject + from header) and normalizes the body through
 * `toContext` (HTML→Markdown when `bodyType === 'html'`, else plain).
 */
export interface MailItem {
  id?: string;
  subject?: string;
  from?: string;
  body: string;
  bodyType?: 'html' | 'text';
}

export function mailItemToContext(item: MailItem, opts: ToContextOptions = {}): ResolvedContext[] {
  const sourceId = item.id ?? 'outlook:item';
  const format = item.bodyType === 'html' ? 'html' : 'plain';
  const text = buildLabelledSource(item);
  return toContext(
    {
      sourceId,
      text,
      format,
      surface: 'outlook',
      ...(item.subject ? { title: item.subject } : {}),
    },
    opts,
  );
}

/**
 * Compose a single labelled source string from the mail item's header fields and body. The
 * `Subject:` / `From:` prefix lines give the engine the same grounding cues a reader sees,
 * and stay legible after the body's HTML→Markdown reduction.
 */
function buildLabelledSource(item: MailItem): string {
  const lines: string[] = [];
  if (item.subject) lines.push(`Subject: ${item.subject}`);
  if (item.from) lines.push(`From: ${item.from}`);
  const header = lines.length > 0 ? `${lines.join('\n')}\n\n` : '';
  return `${header}${item.body}`;
}

/**
 * Reduce a mail body (HTML or text) to plain text **lines** for the `<doc_state>` outline and the
 * body scan. Pure (no DOM): strips tags, decodes the handful of common HTML entities, and splits on
 * block boundaries. This is a lossy, defensive reduction — it is used only to derive an outline/
 * search index, never to render the mail; the full body still flows through `mailItemToContext`'s
 * HTML→Markdown path. Untrusted content stays DATA throughout.
 */
export function mailBodyToLines(body: string, bodyType?: 'html' | 'text'): string[] {
  let text = body;
  if (bodyType === 'html') {
    text = body
      .replace(/<\s*(br|\/p|\/div|\/li|\/h[1-6])\s*\/?\s*>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/gi, "'");
  }
  return text
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter((line) => line.length > 0);
}

/**
 * The active mail item → the `Block[]` the surface-agnostic `buildDocStateSnapshot` consumes for a
 * whole-item `read` (ADR-0006). A mail item has no addressable sub-range, so the "document" is the
 * single active item: the subject becomes the outline heading, and the sender + the leading body
 * lines become paragraph blocks. Bounded by {@link MAX_OUTLINE_LINES} so a long mail can't blow the
 * snapshot budget. The item id is the locator for every block.
 */
export function mailItemToDocStateBlocks(item: MailItem): Block[] {
  const locator = `mail:${item.id ?? 'item'}`;
  const blocks: Block[] = [];
  if (item.subject?.trim()) blocks.push(native.heading(item.subject, 1, locator));
  if (item.from?.trim()) blocks.push(native.paragraph(`From: ${item.from}`, locator));
  for (const line of mailBodyToLines(item.body, item.bodyType).slice(0, MAX_OUTLINE_LINES)) {
    blocks.push(native.paragraph(line, locator));
  }
  return blocks;
}

/** Cap on body lines folded into the whole-item `read` snapshot / the search index. */
export const MAX_OUTLINE_LINES = 40;

/** Cap on matching body lines returned by a lazy `searchDocument` so a common term stays bounded. */
export const MAX_SEARCH_LINES = 8;

/**
 * Scan the mail body for `query` (case-insensitive substring over the body lines) and return the
 * matching lines — labelled with the subject/from header — as context via {@link mailItemToContext}
 * (so the result reuses the same normalization path). Bounded to the first {@link MAX_SEARCH_LINES}
 * matches. Pure: the host read happens in the bridge; this is the match + shaping step. Empty query
 * / no match → `[]`.
 */
export function searchMailItem(item: MailItem, query: string): ResolvedContext[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const matched: string[] = [];
  for (const line of mailBodyToLines(item.body, item.bodyType)) {
    if (line.toLowerCase().includes(needle)) {
      matched.push(line);
      if (matched.length >= MAX_SEARCH_LINES) break;
    }
  }
  if (matched.length === 0) return [];
  return mailItemToContext({
    ...(item.id ? { id: item.id } : {}),
    ...(item.subject ? { subject: item.subject } : {}),
    ...(item.from ? { from: item.from } : {}),
    body: matched.join('\n'),
    bodyType: 'text',
  });
}
