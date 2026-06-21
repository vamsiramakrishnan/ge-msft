import type { ContextKind, ResolvedContext } from '@ge/contracts';
import type { Chunk, ProcessedContent, RawContent, ToContextOptions } from './model.js';
import { toMarkdown } from './normalize.js';
import { parseMarkdownBlocks } from './markdown.js';
import { chunkBlocks } from './chunk.js';
import { contextualizeChunk } from './contextualize.js';

/**
 * The full pipeline: RawContent → Markdown → Block[] → Chunk[].
 * Bridges call this on extracted host/estate content; the chunks then become
 * attach-ready context (see `toContext`).
 */
export function processContent(raw: RawContent, opts: ToContextOptions = {}): ProcessedContent {
  const markdown = toMarkdown(raw);
  const blocks = parseMarkdownBlocks(markdown);
  const chunks = chunkBlocks(blocks, raw, opts);
  return { markdown, blocks, chunks };
}

/**
 * Turn processed content into `ResolvedContext[]` ready to attach to a session.
 *
 * Reference-over-inline: when the source is already indexed in a connected data store and
 * `preferReference` is set, emit a single `indexed-document` reference (ACL-preserving,
 * citations resolve) instead of inlining the chunks. Otherwise emit one contextualized
 * text part per chunk, each carrying its anchor + token estimate for write-back + budgeting.
 */
export function toContext(raw: RawContent, opts: ToContextOptions = {}): ResolvedContext[] {
  if (opts.preferReference && raw.indexedDocumentName) {
    return [
      {
        ref: {
          id: raw.sourceId,
          kind: 'indexed-document',
          surface: raw.surface ?? 'word',
          title: raw.title ?? raw.sourceId,
        },
        value: {
          as: 'indexed-document',
          documentName: raw.indexedDocumentName,
          ...(raw.title ? { title: raw.title } : {}),
        },
      },
    ];
  }
  return processContent(raw, opts).chunks.map((chunk) => chunkToContext(chunk, raw.surface));
}

function chunkToContext(chunk: Chunk, surface?: RawContent['surface']): ResolvedContext {
  const kind: ContextKind = chunk.meta.kinds.includes('table') ? 'table' : 'paragraph';
  const title = chunk.meta.sectionPath.at(-1) ?? chunk.meta.sourceTitle ?? chunk.meta.sourceId;
  const text = contextualizeChunk(chunk);
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
    value: { as: 'text', text, mimeType: 'text/markdown' },
  };
}
