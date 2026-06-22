import type { ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';

/**
 * Pure mapping from OneNote's native object model into grounding-ready context — no Office.js
 * here, so it's unit-testable. The `OneNoteBridge` reads the active page's title and its
 * outlines' rich-text paragraphs via `OneNote.run` and hands the extracted primitives to these
 * functions; they go straight through `@ge/content` (native path) and carry a `page:<id>`
 * write-back locator. OneNote pages have no addressable per-paragraph host id we can write back
 * to safely, so the page id is the locator for every block.
 */

/** An already-extracted OneNote page: title + its rich-text paragraph lines. */
export interface PageElement {
  pageId: string;
  title: string;
  paragraphs: string[];
}

export function pageElementToBlocks(page: PageElement): Block[] {
  const locator = `page:${page.pageId}`;
  const blocks: Block[] = [];
  if (page.title.trim()) blocks.push(native.heading(page.title, 1, locator));
  for (const para of page.paragraphs) {
    if (para.trim()) blocks.push(native.paragraph(para, locator));
  }
  return blocks;
}

/** A OneNote page → attach-ready context, anchored to the page id. */
export function pageToContext(page: PageElement, opts: ToContextOptions = {}): ResolvedContext[] {
  const blocks = pageElementToBlocks(page);
  if (blocks.length === 0) return [];
  const content: NativeContent = {
    sourceId: `on:page:${page.pageId}`,
    surface: 'onenote',
    ...(page.title.trim() ? { title: page.title } : {}),
    blocks,
  };
  return toContextNative(content, opts);
}
