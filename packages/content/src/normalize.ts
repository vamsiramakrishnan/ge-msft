import type { RawContent } from './model.js';

/**
 * Normalize raw content to Markdown — the LLM-native, token-efficient, structure-
 * preserving form (the markitdown/Docling thesis). The heavy format→Markdown lifting
 * (Word OOXML, Excel ranges, PPTX) happens in the bridges; here we cover the common
 * tail: passthrough Markdown, plain-text paragraphs, and a minimal HTML reduction.
 */
export function toMarkdown(raw: RawContent): string {
  switch (raw.format) {
    case 'markdown':
      return raw.text.trim();
    case 'plain':
      return raw.text.trim();
    case 'html':
      return htmlToMarkdown(raw.text).trim();
  }
}

/**
 * Minimal, dependency-free HTML→Markdown reduction for the common structural tags.
 * Not a full converter — bridges should prefer producing Markdown directly — but enough
 * to keep headings/lists/paragraphs/emphasis legible when only HTML is available.
 */
export function htmlToMarkdown(html: string): string {
  let out = html;
  out = out.replace(/<\s*(script|style)[^>]*>[\s\S]*?<\/\s*\1\s*>/gi, '');
  out = out.replace(/<\s*h([1-6])[^>]*>([\s\S]*?)<\/\s*h\1\s*>/gi, (_m, lvl, inner) => {
    return `\n\n${'#'.repeat(Number(lvl))} ${stripTags(inner).trim()}\n\n`;
  });
  out = out.replace(
    /<\s*li[^>]*>([\s\S]*?)<\/\s*li\s*>/gi,
    (_m, inner) => `\n- ${stripTags(inner).trim()}`,
  );
  out = out.replace(/<\s*(p|div)[^>]*>/gi, '\n\n').replace(/<\/\s*(p|div)\s*>/gi, '\n\n');
  out = out.replace(/<\s*br\s*\/?>/gi, '\n');
  out = out.replace(
    /<\s*(strong|b)[^>]*>([\s\S]*?)<\/\s*\1\s*>/gi,
    (_m, _t, inner) => `**${stripTags(inner)}**`,
  );
  out = out.replace(
    /<\s*(em|i)[^>]*>([\s\S]*?)<\/\s*\1\s*>/gi,
    (_m, _t, inner) => `*${stripTags(inner)}*`,
  );
  out = stripTags(out);
  return decodeEntities(out).replace(/\n{3,}/g, '\n\n');
}

function stripTags(s: string): string {
  return s.replace(/<[^>]+>/g, '');
}

function decodeEntities(s: string): string {
  return s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ');
}
