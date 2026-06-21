import type { ResolvedContext } from '@ge/contracts';
import { toContext, type ToContextOptions } from '@ge/content';

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
