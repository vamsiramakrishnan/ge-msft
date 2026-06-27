import { z } from 'zod';
import { asSessionId, type SessionId } from '@ge/contracts';
import { addContextFileUrl, sessionFilesUrl, type GeminiClientConfig } from './config.js';
import { defaultFetch, getJson, postJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

export const DEFAULT_CONTEXT_FILE_MAX_BYTES = 50 * 1024 * 1024;
export const HARD_CONTEXT_FILE_MAX_BYTES = 100 * 1024 * 1024;

const MIME_BY_EXTENSION: Record<string, string> = {
  '.csv': 'text/csv',
  '.tsv': 'text/tab-separated-values',
  '.txt': 'text/plain',
  '.md': 'text/markdown',
  '.json': 'application/json',
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

const EXTENSIONS_BY_MIME = new Map<string, Set<string>>();
for (const [ext, mime] of Object.entries(MIME_BY_EXTENSION)) {
  const set = EXTENSIONS_BY_MIME.get(mime) ?? new Set<string>();
  set.add(ext);
  EXTENSIONS_BY_MIME.set(mime, set);
}

const FALLBACK_MIME_TYPES = new Set(['', 'application/octet-stream', 'binary/octet-stream']);

export interface ContextFileInput {
  /** User-visible file name only. Paths and control characters are rejected. */
  fileName: string;
  /** IANA MIME type. Empty/octet-stream is normalized from the extension when possible. */
  mimeType: string;
  /** Raw file bytes. Text is UTF-8 encoded before upload. */
  contents: string | ArrayBuffer | Uint8Array;
}

export interface ContextFileUploadOptions {
  /**
   * Session id/name to attach the file to. `-` asks Discovery Engine to create a session if needed.
   */
  session?: string;
  /** Per-call byte cap. Defaults to 50 MiB and is always capped at 100 MiB. */
  maxBytes?: number;
  signal?: AbortSignal;
}

export interface NormalizedContextFileInput {
  fileName: string;
  mimeType: string;
  bytes: Uint8Array;
}

export const AddContextFileResponseSchema = z
  .object({
    session: z.string().optional(),
    fileId: z.string().min(1),
    tokenCount: z.union([z.number(), z.string()]).optional(),
  })
  .passthrough();
export type AddContextFileResponse = z.infer<typeof AddContextFileResponseSchema>;

export const ContextFileMetadataSchema = z
  .object({
    fileId: z.string().optional(),
    fileName: z.string().optional(),
    mimeType: z.string().optional(),
    byteSize: z.union([z.number(), z.string()]).optional(),
    tokenCount: z.union([z.number(), z.string()]).optional(),
    downloadUri: z.string().optional(),
  })
  .passthrough();
export type ContextFileMetadata = z.infer<typeof ContextFileMetadataSchema>;

export const ListContextFilesResponseSchema = z
  .object({
    files: z.array(ContextFileMetadataSchema).optional(),
    contextFiles: z.array(ContextFileMetadataSchema).optional(),
    fileMetadata: z.array(ContextFileMetadataSchema).optional(),
  })
  .passthrough();

export interface UploadedContextFile {
  session?: SessionId;
  fileId: string;
  tokenCount?: number;
  fileName: string;
  mimeType: string;
  byteSize: number;
}

export class ContextFileClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  async addContextFile(
    input: ContextFileInput,
    opts: ContextFileUploadOptions = {},
  ): Promise<UploadedContextFile> {
    const normalized = normalizeContextFileInput(input, opts);
    const session = opts.session ?? '-';
    const raw = await postJson(
      addContextFileUrl(this.config, session),
      {
        fileName: normalized.fileName,
        mimeType: normalized.mimeType,
        fileContents: bytesToBase64(normalized.bytes),
      },
      this.tokens,
      this.fetchImpl,
      opts.signal,
      { maxAttempts: 1 }, // uploads are not retried; duplicate files are not provably idempotent.
    );
    const parsed = AddContextFileResponseSchema.parse(raw);
    return {
      ...(parsed.session ? { session: asSessionId(parsed.session) } : {}),
      fileId: parsed.fileId,
      ...(parsed.tokenCount !== undefined ? { tokenCount: Number(parsed.tokenCount) } : {}),
      fileName: normalized.fileName,
      mimeType: normalized.mimeType,
      byteSize: normalized.bytes.byteLength,
    };
  }

  async listContextFiles(
    session: string,
    opts: { signal?: AbortSignal } = {},
  ): Promise<{
    files: ContextFileMetadata[];
  }> {
    const raw = await getJson(
      sessionFilesUrl(this.config, session),
      this.tokens,
      this.fetchImpl,
      opts.signal,
    );
    const parsed = ListContextFilesResponseSchema.parse(raw);
    return { files: parsed.files ?? parsed.contextFiles ?? parsed.fileMetadata ?? [] };
  }
}

export function normalizeContextFileInput(
  input: ContextFileInput,
  opts: Pick<ContextFileUploadOptions, 'maxBytes'> = {},
): NormalizedContextFileInput {
  const fileName = normalizeFileName(input.fileName);
  const ext = extensionOf(fileName);
  const mimeType = normalizeMimeType(input.mimeType, ext);
  assertExtensionMatchesMime(fileName, ext, mimeType);

  const bytes = toBytes(input.contents);
  const maxBytes = Math.min(
    opts.maxBytes ?? DEFAULT_CONTEXT_FILE_MAX_BYTES,
    HARD_CONTEXT_FILE_MAX_BYTES,
  );
  if (bytes.byteLength === 0) throw new Error('Context file upload rejected: file is empty.');
  if (bytes.byteLength > maxBytes) {
    throw new Error(
      `Context file upload rejected: ${bytes.byteLength} bytes exceeds the ${maxBytes} byte limit.`,
    );
  }
  return { fileName, mimeType, bytes };
}

export function supportedContextFileFormats(): Array<{ extension: string; mimeType: string }> {
  return Object.entries(MIME_BY_EXTENSION).map(([extension, mimeType]) => ({
    extension,
    mimeType,
  }));
}

export function bytesToBase64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  let out = '';
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i]! << 16) | (bytes[i + 1]! << 8) | bytes[i + 2]!;
    out += alphabet[(n >> 18) & 63]!;
    out += alphabet[(n >> 12) & 63]!;
    out += alphabet[(n >> 6) & 63]!;
    out += alphabet[n & 63]!;
  }
  if (i < bytes.length) {
    const a = bytes[i]!;
    const b = i + 1 < bytes.length ? bytes[i + 1]! : 0;
    const n = (a << 16) | (b << 8);
    out += alphabet[(n >> 18) & 63]!;
    out += alphabet[(n >> 12) & 63]!;
    out += i + 1 < bytes.length ? alphabet[(n >> 6) & 63]! : '=';
    out += '=';
  }
  return out;
}

function normalizeFileName(fileName: string): string {
  const trimmed = fileName.trim();
  if (!trimmed) throw new Error('Context file upload rejected: fileName is required.');
  if (trimmed.length > 256) {
    throw new Error('Context file upload rejected: fileName must be 256 characters or less.');
  }
  if (trimmed === '.' || trimmed === '..' || hasUnsafeFileNameChar(trimmed)) {
    throw new Error('Context file upload rejected: fileName must be a plain file name.');
  }
  return trimmed;
}

function hasUnsafeFileNameChar(fileName: string): boolean {
  for (const ch of fileName) {
    const code = ch.charCodeAt(0);
    if (ch === '/' || ch === '\\' || code <= 0x1f || code === 0x7f) return true;
  }
  return false;
}

function extensionOf(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot >= 0 ? fileName.slice(dot).toLowerCase() : '';
}

function normalizeMimeType(mimeType: string, ext: string): string {
  const lower = mimeType.trim().toLowerCase();
  if (FALLBACK_MIME_TYPES.has(lower)) {
    const inferred = MIME_BY_EXTENSION[ext];
    if (!inferred) {
      throw new Error(
        'Context file upload rejected: MIME type is required for files without a supported extension.',
      );
    }
    return inferred;
  }
  if (!EXTENSIONS_BY_MIME.has(lower)) {
    throw new Error(
      `Context file upload rejected: unsupported MIME type '${mimeType}'. Supported extensions: ${Object.keys(
        MIME_BY_EXTENSION,
      )
        .sort()
        .join(', ')}.`,
    );
  }
  return lower;
}

function assertExtensionMatchesMime(fileName: string, ext: string, mimeType: string): void {
  const allowed = EXTENSIONS_BY_MIME.get(mimeType);
  if (!allowed) return;
  if (!allowed.has(ext)) {
    throw new Error(
      `Context file upload rejected: '${fileName}' extension does not match MIME type '${mimeType}'.`,
    );
  }
}

function toBytes(contents: ContextFileInput['contents']): Uint8Array {
  if (typeof contents === 'string') return new TextEncoder().encode(contents);
  if (contents instanceof Uint8Array) return contents;
  return new Uint8Array(contents);
}
