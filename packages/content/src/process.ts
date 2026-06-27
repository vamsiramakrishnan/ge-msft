import type { ContextKind, ResolvedContext, Surface } from '@ge/contracts';
import type {
  Chunk,
  NativeContent,
  ProcessedContent,
  RawContent,
  SourceMeta,
  ToContextOptions,
} from './model.js';
import { toMarkdown } from './normalize.js';
import { parseMarkdownBlocks } from './markdown.js';
import { chunkBlocks } from './chunk.js';
import { contextualizeChunk } from './contextualize.js';

/**
 * Native path (preferred): bridges that read the Office object model pass `Block[]`
 * directly — no Markdown round-trip, native locators preserved.
 */
export function processNative(
  native: NativeContent,
  opts: ToContextOptions = {},
): ProcessedContent {
  const chunks = chunkBlocks(native.blocks, native, opts);
  return { blocks: native.blocks, chunks };
}

/**
 * String path (fallback): for sources that arrive as text/HTML (e.g. an Outlook body).
 * Normalize to Markdown, then parse into blocks with char offsets.
 */
export function processContent(raw: RawContent, opts: ToContextOptions = {}): ProcessedContent {
  const markdown = toMarkdown(raw);
  const blocks = parseMarkdownBlocks(markdown);
  const chunks = chunkBlocks(blocks, raw, opts);
  return { blocks, chunks };
}

/** Native content → attach-ready `ResolvedContext[]`. */
export function toContextNative(
  native: NativeContent,
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const ref = referenceOnly(native, opts);
  if (ref) return ref;
  return processNative(native, opts).chunks.map((c) => chunkToContext(c, native.surface));
}

/** String content → attach-ready `ResolvedContext[]`. */
export function toContext(raw: RawContent, opts: ToContextOptions = {}): ResolvedContext[] {
  const ref = referenceOnly(raw, opts);
  if (ref) return ref;
  return processContent(raw, opts).chunks.map((c) => chunkToContext(c, raw.surface));
}

/**
 * Reference-over-inline: when the source is already indexed in a connected data store and
 * the caller prefers it, emit a single `indexed-document` reference (ACL-preserving,
 * citations resolve) instead of inlining content.
 */
function referenceOnly(meta: SourceMeta, opts: ToContextOptions): ResolvedContext[] | null {
  if (!opts.preferReference || !meta.indexedDocumentName) return null;
  return [
    {
      ref: {
        id: meta.sourceId,
        kind: 'indexed-document',
        surface: meta.surface ?? 'word',
        title: meta.title ?? meta.sourceId,
      },
      value: {
        as: 'indexed-document',
        documentName: meta.indexedDocumentName,
        ...(meta.title ? { title: meta.title } : {}),
      },
    },
  ];
}

function chunkToContext(chunk: Chunk, surface?: Surface): ResolvedContext {
  const kind: ContextKind = chunk.meta.kinds.includes('table') ? 'table' : 'paragraph';
  const title = chunk.meta.sectionPath.at(-1) ?? chunk.meta.sourceTitle ?? chunk.meta.sourceId;
  return {
    ref: {
      id: chunk.id,
      kind,
      surface: surface ?? chunk.meta.surface ?? 'word',
      title,
      preview: chunk.text.slice(0, 120),
      tokensEstimate: chunk.meta.tokensEstimate,
      anchor: chunk.meta.anchor,
    },
    value: { as: 'text', text: contextualizeChunk(chunk), mimeType: 'text/markdown' },
  };
}
