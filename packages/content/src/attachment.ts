import type { ResolvedContext, Surface } from '@ge/contracts';
import { toContext } from './process.js';
import type { ToContextOptions } from './model.js';

/**
 * Attachments as context — the shared pipeline for a mail attachment (Outlook) or a file shared in
 * a chat/meeting (Teams via Graph). Surface-agnostic: the bridge or estate reader hands in the
 * file's name, MIME type and (when read) its bytes; this module decides how it grounds.
 *
 *   • text-like files (plain text, Markdown, CSV/TSV, JSON, XML, HTML, .eml, .ics) are decoded and
 *     normalized through the same `toContext` path as a mail body — bounded, chunked, DATA not
 *     instructions;
 *   • engine-ingestible binaries (PDF, Word, Excel, PowerPoint) become ONE `file` part marked
 *     `uploadable`: the runtime uploads the bytes as a session context file on an explicit user
 *     attach and records the `fileId`; nothing is parsed locally, nothing is inlined;
 *   • everything else (images, archives, executables…) becomes a `file` part the model can see by
 *     name only, with a `note` saying why the content is unavailable.
 */

export type AttachmentClass = 'text' | 'file' | 'unsupported';

/** Inline decode ceiling for text-like attachments; larger ones ground by name only. */
export const MAX_INLINE_ATTACHMENT_BYTES = 2 * 1024 * 1024;

const TEXT_MIME_BY_EXTENSION: Record<string, string> = {
  '.txt': 'text/plain',
  '.log': 'text/plain',
  '.md': 'text/markdown',
  '.markdown': 'text/markdown',
  '.csv': 'text/csv',
  '.tsv': 'text/tab-separated-values',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.eml': 'message/rfc822',
  '.ics': 'text/calendar',
};

/** Binary formats Discovery Engine ingests as session context files (see gemini-client/context-files). */
const UPLOADABLE_MIME_BY_EXTENSION: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
};

const TEXT_MIMES = new Set(Object.values(TEXT_MIME_BY_EXTENSION));
const UPLOADABLE_MIMES = new Set(Object.values(UPLOADABLE_MIME_BY_EXTENSION));
const OPAQUE_MIMES = new Set(['', 'application/octet-stream', 'binary/octet-stream']);

function extensionOf(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot === -1 ? '' : fileName.slice(dot).toLowerCase();
}

/**
 * Decide how an attachment grounds from its name and declared MIME type. The extension wins when
 * the declared type is missing or opaque (`application/octet-stream`), which is what Outlook and
 * Graph report for many Office files; `text/*` not in the table still counts as text.
 */
export function classifyAttachment(
  fileName: string,
  mimeType?: string,
): { cls: AttachmentClass; mimeType: string } {
  const declared = (mimeType ?? '').split(';')[0]!.trim().toLowerCase();
  const ext = extensionOf(fileName);
  // A known extension wins: hosts mislabel Office files as octet-stream and text/plain alike,
  // and a `.docx` declared `text/plain` would otherwise be decoded as text and fail.
  const byExtText = TEXT_MIME_BY_EXTENSION[ext];
  if (byExtText) return { cls: 'text', mimeType: byExtText };
  const byExtUpload = UPLOADABLE_MIME_BY_EXTENSION[ext];
  if (byExtUpload) return { cls: 'file', mimeType: byExtUpload };
  if (OPAQUE_MIMES.has(declared)) {
    return { cls: 'unsupported', mimeType: declared || 'application/octet-stream' };
  }
  if (TEXT_MIMES.has(declared) || declared.startsWith('text/')) {
    return { cls: 'text', mimeType: declared };
  }
  if (UPLOADABLE_MIMES.has(declared)) return { cls: 'file', mimeType: declared };
  return { cls: 'unsupported', mimeType: declared };
}

export interface AttachmentInput {
  /** Stable context id for the attachment (e.g. `outlook:attachment:<id>`, `teams:file:<id>`). */
  sourceId: string;
  surface: Surface;
  fileName: string;
  mimeType?: string;
  sizeBytes?: number;
  /** Decoded bytes, when the caller fetched them. Required to inline a text-like attachment. */
  bytes?: Uint8Array;
  /** Already-decoded text (an `.eml` body Outlook returns as a string). Takes precedence over bytes. */
  text?: string;
}

/**
 * Turn one attachment into attach-ready context. Text-like content goes through `toContext`
 * (HTML→Markdown, chunked) with a `File:` header so the engine sees which file each chunk came from;
 * anything else is a single `file` part. Never throws on undecodable bytes: a text-like file that
 * fails to decode degrades to a by-name `file` part with a note.
 */
export function attachmentToContext(
  input: AttachmentInput,
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const { cls, mimeType } = classifyAttachment(input.fileName, input.mimeType);
  const size = input.sizeBytes ?? input.bytes?.byteLength;
  const ref = {
    id: input.sourceId,
    kind: 'attachment' as const,
    surface: input.surface,
    title: input.fileName,
    mimeType,
    ...(size !== undefined ? { sizeBytes: size } : {}),
  };
  const filePart = (uploadable: boolean, note?: string): ResolvedContext => ({
    ref,
    value: {
      as: 'file',
      fileName: input.fileName,
      mimeType,
      ...(size !== undefined ? { sizeBytes: size } : {}),
      uploadable,
      ...(note ? { note } : {}),
    },
  });

  if (cls === 'file') return [filePart(true)];
  if (cls === 'unsupported') {
    return [
      filePart(false, `${mimeType || 'this file type'} is not readable here; content omitted`),
    ];
  }

  // Bound before decoding: the declared size, or the bytes actually handed in.
  if (size !== undefined && size > MAX_INLINE_ATTACHMENT_BYTES) {
    return [filePart(false, 'too large to inline; content omitted')];
  }
  const text = input.text ?? decodeText(input.bytes);
  if (text === undefined) return [filePart(false, 'content was not read; attach it to read it')];
  const format =
    mimeType === 'text/html' ? 'html' : mimeType === 'text/markdown' ? 'markdown' : 'plain';
  const parts = toContext(
    {
      sourceId: input.sourceId,
      text: `File: ${input.fileName}\n\n${text}`,
      format,
      surface: input.surface,
      title: input.fileName,
    },
    opts,
  );
  // Chunks inherit the attachment identity so the pane shows which file they belong to.
  return parts.map((part) => ({
    ...part,
    ref: { ...part.ref, kind: 'attachment' as const, mimeType },
  }));
}

/** UTF-8 decode (BOM-tolerant). `undefined` when there are no bytes or they are not valid text. */
export function decodeText(bytes: Uint8Array | undefined): string | undefined {
  if (!bytes) return undefined;
  try {
    const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  } catch {
    return undefined;
  }
}

/** Base64 (standard or URL-safe) → bytes, without Node's Buffer so it runs in the task pane. */
export function decodeBase64(data: string): Uint8Array {
  const normalized = data.replace(/-/g, '+').replace(/_/g, '/').replace(/\s+/g, '');
  const padded = normalized + '='.repeat((4 - (normalized.length % 4)) % 4);
  const binary = atob(padded);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) out[i] = binary.charCodeAt(i);
  return out;
}
