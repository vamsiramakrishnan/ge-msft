/**
 * @ge/content — process host/estate content into grounding-ready context.
 *
 * Pipeline (OSS-inspired): normalize to Markdown (markitdown/Docling) → structure-aware
 * recursive chunking (LangChain) with section grouping (LlamaIndex parent/child) →
 * Anthropic-style contextualization → metadata-rich, anchored `ResolvedContext[]` that
 * plug straight into @ge/gemini-client's SessionContext / query.parts[].
 * See docs/CONTENT-PROCESSING.md.
 */
export * from './model.js';
export { estimateTokens } from './tokens.js';
export { toMarkdown, htmlToMarkdown } from './normalize.js';
export { parseMarkdownBlocks, tableToMarkdown } from './markdown.js';
export { chunkBlocks, splitText } from './chunk.js';
export { contextualizeChunk } from './contextualize.js';
export { processContent, toContext } from './process.js';
