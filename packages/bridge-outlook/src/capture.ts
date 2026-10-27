import type { Block } from '@ge/content';
import type { ContextRef, ResolvedContext } from '@ge/contracts';
import {
  attachmentToContext,
  classifyAttachment,
  decodeBase64,
  native,
  toContext,
  type AttachmentInput,
  type ToContextOptions,
} from '@ge/content';

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

/* ───────────────────────────── attachments ───────────────────────────── */

/**
 * The read-mode view of one mail attachment (`item.attachments[]`, Mailbox 1.1). Inline images
 * (`isInline`) are part of the rendered body, not grounding; `item` attachments are forwarded
 * messages the host returns as `.eml` text; `cloud` attachments are links the host returns as a URL.
 */
export interface MailAttachment {
  id: string;
  name: string;
  attachmentType?: 'file' | 'item' | 'cloud' | string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
}

/** Stable context id for a mail attachment: scoped to the item so a different mail never aliases it. */
export function attachmentRefId(itemId: string | undefined, attachmentId: string): string {
  return `outlook:attachment:${itemId ?? 'item'}:${attachmentId}`;
}

/**
 * The attachable handles for a mail item's attachments: one `attachment` ref per non-inline file,
 * forwarded item, or cloud link, with the host pointer the bridge needs to fetch its content. Pure:
 * no content is read here, so listing stays cheap (`attachments` / `list attachment`).
 */
export function mailAttachmentsToRefs(
  itemId: string | undefined,
  attachments: readonly MailAttachment[],
): ContextRef[] {
  return attachments
    .filter((a) => !a.isInline && a.id && a.name)
    .map((a): ContextRef => {
      const { cls, mimeType } = classifyAttachment(a.name, a.contentType);
      const kindNote =
        a.attachmentType === 'item'
          ? 'forwarded message'
          : a.attachmentType === 'cloud'
            ? 'cloud link'
            : cls === 'text'
              ? 'readable text'
              : cls === 'file'
                ? 'attach to read as a context file'
                : 'content not readable';
      return {
        id: attachmentRefId(itemId, a.id),
        kind: 'attachment',
        surface: 'outlook',
        title: a.name,
        preview: `${kindNote}${a.size ? ` · ${formatSize(a.size)}` : ''}`,
        mimeType,
        ...(a.size !== undefined ? { sizeBytes: a.size } : {}),
        hostRef: { type: 'outlook.attachment', itemId: itemId ?? 'item', attachmentId: a.id },
      };
    });
}

/** What `getAttachmentContentAsync` returned for one attachment (Mailbox 1.8), host-format tagged. */
export interface MailAttachmentContent {
  /** `base64` for files, `eml` for forwarded messages, `url` for cloud links, `iCalendar` for invites. */
  format: 'base64' | 'eml' | 'url' | 'iCalendar' | string;
  content: string;
}

/**
 * Turn one attachment (plus its fetched content, when available) into attach-ready context through
 * the shared `attachmentToContext` pipeline. Forwarded messages and invites are text; a cloud link
 * grounds as a one-line reference (the file itself lives in OneDrive/SharePoint); a `base64` file
 * is decoded and classified. Without content, the attachment grounds by name with a note.
 */
export function mailAttachmentToContext(
  itemId: string | undefined,
  attachment: MailAttachment,
  content?: MailAttachmentContent,
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const sourceId = attachmentRefId(itemId, attachment.id);
  const base = {
    sourceId,
    surface: 'outlook' as const,
    fileName: attachment.name,
    ...(attachment.contentType ? { mimeType: attachment.contentType } : {}),
    ...(attachment.size !== undefined ? { sizeBytes: attachment.size } : {}),
  };
  const hostRef = {
    type: 'outlook.attachment' as const,
    itemId: itemId ?? 'item',
    attachmentId: attachment.id,
  };
  // Every resolved part keeps the host pointer so a later byte read can find the attachment.
  return mailAttachmentParts(base, sourceId, attachment, content, opts).map((part) => ({
    ...part,
    ref: { ...part.ref, hostRef },
  }));
}

function mailAttachmentParts(
  base: Omit<AttachmentInput, 'bytes' | 'text'>,
  sourceId: string,
  attachment: MailAttachment,
  content: MailAttachmentContent | undefined,
  opts: ToContextOptions,
): ResolvedContext[] {
  if (!content) return attachmentToContext(base, opts);
  switch (content.format) {
    case 'eml':
      return attachmentToContext(
        {
          ...base,
          fileName: ensureExtension(attachment.name, '.eml'),
          mimeType: 'message/rfc822',
          text: content.content,
        },
        opts,
      );
    case 'iCalendar':
      return attachmentToContext(
        {
          ...base,
          fileName: ensureExtension(attachment.name, '.ics'),
          mimeType: 'text/calendar',
          text: content.content,
        },
        opts,
      );
    case 'url':
      return toContext(
        {
          sourceId,
          text: `File: ${attachment.name}\n\nCloud attachment link: ${content.content}`,
          format: 'plain',
          surface: 'outlook',
          title: attachment.name,
        },
        opts,
      ).map((part) => ({ ...part, ref: { ...part.ref, kind: 'attachment' as const } }));
    default: {
      let bytes: Uint8Array | undefined;
      try {
        bytes = decodeBase64(content.content);
      } catch {
        bytes = undefined;
      }
      return attachmentToContext({ ...base, ...(bytes ? { bytes } : {}) }, opts);
    }
  }
}

function ensureExtension(name: string, ext: string): string {
  return name.toLowerCase().endsWith(ext) ? name : `${name}${ext}`;
}

function formatSize(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${Math.round(n / 1024)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}
