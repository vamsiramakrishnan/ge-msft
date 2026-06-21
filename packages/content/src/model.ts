import type { Anchor, ResolvedContext, Surface } from '@ge/contracts';

/**
 * The content-processing model. Pipeline (inspired by markitdown/Docling →
 * LangChain recursive + structure-aware splitting → LlamaIndex parent/child →
 * Anthropic contextual retrieval):
 *
 *   RawContent → normalize to Markdown → Block[] (structure + char offsets)
 *             → Chunk[] (token-budgeted, section-aware, with write-back anchors)
 *             → contextualized ResolvedContext[] (ready for query.parts[]).
 */

/** What a bridge hands in: extracted host/estate content + where it came from. */
export interface RawContent {
  sourceId: string; // stable id for this source (e.g. "word:body", "sp:contract.docx")
  text: string;
  format: 'markdown' | 'plain' | 'html';
  title?: string;
  surface?: Surface;
  /** If the source is already indexed in a connected data store, its VAIS doc name. */
  indexedDocumentName?: string;
}

export type BlockKind = 'heading' | 'paragraph' | 'list' | 'table' | 'code' | 'quote';

/** A structural unit of the normalized Markdown, with char offsets into the source. */
export interface Block {
  kind: BlockKind;
  level?: number; // heading level 1..6
  text: string;
  start: number; // inclusive char offset
  end: number; // exclusive char offset
}

export interface ChunkMeta {
  sourceId: string;
  sourceTitle?: string;
  surface?: Surface;
  sectionPath: string[]; // heading breadcrumb, e.g. ["5. Service Levels", "5.2 Availability"]
  kinds: BlockKind[]; // which block kinds the chunk contains
  charStart: number;
  charEnd: number;
  tokensEstimate: number;
  anchor: Anchor; // content anchor for write-back
}

export interface Chunk {
  id: string; // `${sourceId}#${index}`
  index: number;
  text: string; // the chunk body (Markdown), before contextualization
  meta: ChunkMeta;
}

export interface ProcessedContent {
  markdown: string;
  blocks: Block[];
  chunks: Chunk[];
}

export interface ChunkOptions {
  /** Target max tokens per chunk (soft). Default 400. */
  maxTokens?: number;
  /** Sentence overlap tokens carried between split chunks of one oversized block. Default 40. */
  overlapTokens?: number;
  /** Start a new chunk at headings of this level or shallower. Default 2. */
  sectionBreakLevel?: number;
}

export interface ToContextOptions extends ChunkOptions {
  /** Prefer an indexed-document reference over inline text when available (ACL-preserving). */
  preferReference?: boolean;
}

/** Convenience alias for the attach-ready output. */
export type ContextBundle = ResolvedContext[];
