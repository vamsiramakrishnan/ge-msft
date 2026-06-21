import type { Chunk } from './model.js';

/**
 * Anthropic-style **contextual retrieval**: prepend a compact, self-situating header to
 * each chunk so it grounds well out of context (source + section breadcrumb + position).
 * Kept to a single terse line so it costs a handful of tokens, not 50–100 — the engine
 * already has the citation machinery; we just need the chunk to know where it lives.
 */
export function contextualizeChunk(chunk: Chunk): string {
  const { sourceTitle, sectionPath } = chunk.meta;
  const crumbs = [sourceTitle, ...sectionPath].filter(Boolean).join(' › ');
  const header = crumbs ? `[${crumbs}]` : '';
  return header ? `${header}\n${chunk.text}` : chunk.text;
}
