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

/**
 * A OneNote page → the `Block[]` the surface-agnostic `buildDocStateSnapshot` consumes for the
 * `<doc_state>` outline/inventory (ADR-0003 Layer B element 1). Reuses {@link pageElementToBlocks}
 * so the snapshot's outline (the page title heading) + inventory come from the SAME native mapping
 * as grounding context.
 */
export function pageElementToDocStateBlocks(page: PageElement): Block[] {
  return pageElementToBlocks(page);
}

/** Cap on paragraphs returned by a lazy `searchDocument` so a common term can't blow the budget. */
export const MAX_SEARCH_PARAGRAPHS = 8;

/**
 * Scan a page's paragraphs for `query` (case-insensitive substring) and return the matching
 * paragraphs — still anchored to the page id — as context via {@link pageToContext}, bounded to the
 * first {@link MAX_SEARCH_PARAGRAPHS} matches. Pure: the host read happens in the bridge; this is
 * the match + shaping step. The title is preserved so a hit keeps its page heading. Empty query /
 * no match → `[]`.
 */
export function searchPage(page: PageElement, query: string): ResolvedContext[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const matched: string[] = [];
  for (const para of page.paragraphs) {
    if (para.toLowerCase().includes(needle)) {
      matched.push(para);
      if (matched.length >= MAX_SEARCH_PARAGRAPHS) break;
    }
  }
  if (matched.length === 0) return [];
  return pageToContext({ pageId: page.pageId, title: page.title, paragraphs: matched });
}
