import type { Anchor, ResolvedContext, Surface } from '@ge/contracts';

/**
 * The content-processing model. Two ways in, one pipeline:
 *
 *   (native, preferred)  Office object model → Block[] → Chunk[] → ResolvedContext[]
 *   (string fallback)    text/HTML → Markdown → Block[] → Chunk[] → ResolvedContext[]
 *
 * The Office host already exposes structure (Word paragraphs/tables/content-controls,
 * Excel ranges, PowerPoint slides) — so bridges build `Block[]` directly and skip the
 * Markdown regex reparse. The string path is only for sources that genuinely arrive as
 * a blob of text/HTML (e.g. an Outlook message body). This is NOT RAG: no embeddings,
 * no vector store, no retrieval — the engine grounds; we only normalize + label what we
 * attach as `query.parts[]`.
 */

/** Native structured data (e.g. an Excel range): the Office model, not re-parsed text. */
export interface StructuredData {
  columns: string[];
  rows: (string | number)[][];
}

/** Identity + provenance of a source the content came from. */
export interface SourceMeta {
  sourceId: string; // stable id (e.g. "word:body", "xl:Sheet1!A1:D9", "sp:contract.docx")
  title?: string;
  surface?: Surface;
  /** If the source is already indexed in a connected data store, its VAIS doc name. */
  indexedDocumentName?: string;
}

/** What a bridge hands in when it only has a string (Outlook body, plain text). */
export interface RawContent extends SourceMeta {
  text: string;
  format: 'markdown' | 'plain' | 'html';
}

/** What a bridge hands in when it has native structure (Word/Excel/PowerPoint object model). */
export interface NativeContent extends SourceMeta {
  blocks: Block[];
}

export type BlockKind = 'heading' | 'paragraph' | 'list' | 'table' | 'code' | 'quote';

/**
 * A structural unit. From the string path it carries char offsets (`start`/`end`); from
 * the native path it carries a host `locator` (content-control id, range address, slide
 * index) — whichever the source can provide for write-back.
 */
export interface Block {
  kind: BlockKind;
  level?: number; // heading level 1..6 (or PowerPoint: derived)
  text: string; // the block as Markdown (tables already GFM)
  start?: number; // char offset (string path)
  end?: number;
  locator?: string; // native host handle, e.g. "cc:42", "range:Sheet1!A1:D9", "slide:4"
  data?: StructuredData; // for kind 'table' from the native path
}

export interface ChunkMeta {
  sourceId: string;
  sourceTitle?: string;
  surface?: Surface;
  sectionPath: string[]; // heading breadcrumb, e.g. ["5. Service Levels", "5.2 Availability"]
  kinds: BlockKind[];
  charStart?: number;
  charEnd?: number;
  tokensEstimate: number;
  anchor: Anchor; // content anchor (+ native locator) for write-back
}

export interface Chunk {
  id: string; // `${sourceId}#${index}`
  index: number;
  text: string; // the chunk body (Markdown), before contextualization
  meta: ChunkMeta;
}

export interface ProcessedContent {
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
