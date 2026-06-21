/**
 * @ge/content — process host/estate content into grounding-ready context.
 *
 * NOT RAG: no embeddings, no vector store, no retrieval. Gemini Enterprise grounds; this
 * package only normalizes + chunks + labels what the bridges attach as query.parts[].
 *
 * Two ways in, one pipeline:
 *   • native (preferred): Office object model → Block[] (via `native` builders)
 *   • string (fallback):  text/HTML → Markdown → Block[]
 * Both → structure-aware, section-grouped, token-budgeted, anchored Chunk[] →
 * ResolvedContext[] that plug into @ge/gemini-client's SessionContext / query.parts[].
 * See docs/CONTENT-PROCESSING.md.
 */
export * from './model.js';
export * as native from './native.js';
export { estimateTokens } from './tokens.js';
export { toMarkdown, htmlToMarkdown } from './normalize.js';
export { parseMarkdownBlocks, tableToMarkdown } from './markdown.js';
export { chunkBlocks, splitText } from './chunk.js';
export { contextualizeChunk } from './contextualize.js';
export { processContent, processNative, toContext, toContextNative } from './process.js';
export {
  recommendStrategy,
  ContextBudget,
  DEFAULT_MAX_INLINE_TOKENS,
  DEFAULT_CONTEXT_BUDGET_TOKENS,
  type AttachStrategy,
  type StrategyInput,
} from './budget.js';
