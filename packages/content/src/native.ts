import type { Block, StructuredData } from './model.js';
import { tableToMarkdown } from './markdown.js';

/**
 * Block builders for bridges that read the Office object model directly. Using these
 * avoids the string→Markdown→regex round-trip: a Word paragraph, an Excel range, or a
 * PowerPoint slide becomes a `Block` with a **native host locator** for write-back.
 */

export function heading(text: string, level: number, locator?: string): Block {
  return {
    kind: 'heading',
    level,
    text: `${'#'.repeat(Math.min(Math.max(level, 1), 6))} ${text}`,
    ...(locator ? { locator } : {}),
  };
}

export function paragraph(text: string, locator?: string): Block {
  return { kind: 'paragraph', text, ...(locator ? { locator } : {}) };
}

export function listBlock(items: string[], locator?: string): Block {
  return {
    kind: 'list',
    text: items.map((i) => `- ${i}`).join('\n'),
    ...(locator ? { locator } : {}),
  };
}

export function quote(text: string, locator?: string): Block {
  return { kind: 'quote', text: text.replace(/^/gm, '> '), ...(locator ? { locator } : {}) };
}

export function code(text: string, locator?: string): Block {
  return { kind: 'code', text: '```\n' + text + '\n```', ...(locator ? { locator } : {}) };
}

/** An Excel range (or any tabular host object) as a native table block. */
export function table(data: StructuredData, locator?: string): Block {
  return {
    kind: 'table',
    text: tableToMarkdown(data.columns, data.rows),
    data,
    ...(locator ? { locator } : {}),
  };
}

/** A PowerPoint slide as a heading (title) + its body text, anchored to the slide index. */
export function slide(index: number, title: string, body: string[], slideId?: string): Block[] {
  const locator = slideId ? `slide:${slideId}` : `slide:${index}`;
  const blocks: Block[] = [heading(title || `Slide ${index + 1}`, 2, locator)];
  for (const para of body) if (para.trim()) blocks.push(paragraph(para, locator));
  return blocks;
}
