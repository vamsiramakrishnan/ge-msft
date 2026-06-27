import type { ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';

/**
 * Pure mapping from PowerPoint's native object model into grounding-ready context — no
 * Office.js here, so it's unit-testable. The `PowerPointBridge` reads selected slides (their
 * shapes' text + speaker notes) via `PowerPoint.run` and hands the extracted primitives to
 * these functions; they go straight through `@ge/content` (native path, no Markdown
 * round-trip) using the `native.slide()` builder and carry a `slide:<id>` write-back locator.
 */

/** An already-extracted PowerPoint slide: title-ish first line, body lines, optional notes. */
export interface SlideElement {
  /** Zero-based slide index (position in the deck). */
  index: number;
  /** Stable host slide id, used as the write-back locator when present. */
  slideId?: string;
  /** The slide title (first shape's text, by convention) — may be empty. */
  title: string;
  /** The remaining shape text lines on the slide. */
  body: string[];
  /** Speaker notes for the slide, if any. */
  notes?: string;
}

/**
 * Split a slide's shape texts into a title (first non-empty) + body lines. The host hands us
 * the raw per-shape text; the first non-empty line is treated as the title, the rest as body.
 * A multi-line shape contributes one body line per non-empty line.
 */
export function shapesToSlideText(shapeTexts: string[]): { title: string; body: string[] } {
  const lines: string[] = [];
  for (const raw of shapeTexts) {
    for (const line of raw.split(/\r?\n/)) {
      if (line.trim().length > 0) lines.push(line.trim());
    }
  }
  const title = lines[0] ?? '';
  return { title, body: lines.slice(1) };
}

/** Turn captured slides into native blocks via the `native.slide()` builder. */
export function slideElementsToBlocks(slides: SlideElement[]): Block[] {
  const blocks: Block[] = [];
  for (const s of slides) {
    const body = s.notes && s.notes.trim() ? [...s.body, `Notes: ${s.notes.trim()}`] : s.body;
    blocks.push(...native.slide(s.index, s.title, body, s.slideId));
  }
  return blocks;
}

/** Captured slides → attach-ready context, anchored per-slide. */
export function slidesToContext(
  sourceId: string,
  title: string | undefined,
  slides: SlideElement[],
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const blocks = slideElementsToBlocks(slides);
  if (blocks.length === 0) return [];
  const content: NativeContent = {
    sourceId,
    surface: 'powerpoint',
    ...(title ? { title } : {}),
    blocks,
  };
  return toContextNative(content, opts);
}

/** A single selected slide → context (same mapping as `slidesToContext`). */
export function selectedSlideToContext(slide: SlideElement): ResolvedContext[] {
  return slidesToContext(`pp:slide:${slide.slideId ?? slide.index}`, undefined, [slide]);
}

/**
 * Captured slides → the `Block[]` the surface-agnostic `buildDocStateSnapshot` consumes for the
 * `<doc_state>` outline/inventory (ADR-0003 Layer B element 1). Reuses {@link slideElementsToBlocks}
 * so the snapshot's slide inventory comes from the SAME native mapping as grounding context — each
 * slide's title becomes a `slide:<id>` heading the builder lists under `inventory`.
 */
export function slideElementsToDocStateBlocks(slides: SlideElement[]): Block[] {
  return slideElementsToBlocks(slides);
}

/** Cap on slides scanned/returned by a lazy `searchDocument` so a common term can't blow the budget. */
export const MAX_SEARCH_SLIDES = 8;

/**
 * Scan captured slides for `query` (case-insensitive substring over the title + body lines) and
 * return the matching slides as context via {@link slidesToContext}, bounded to the first
 * {@link MAX_SEARCH_SLIDES} matches. Pure: the host read happens in the bridge; this is the match +
 * shaping step. Empty query / no match → `[]`.
 */
export function searchSlides(slides: SlideElement[], query: string): ResolvedContext[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const matched: SlideElement[] = [];
  for (const slide of slides) {
    const haystack = [slide.title, ...slide.body].join('\n').toLowerCase();
    if (haystack.includes(needle)) {
      matched.push(slide);
      if (matched.length >= MAX_SEARCH_SLIDES) break;
    }
  }
  if (matched.length === 0) return [];
  return slidesToContext('pp:search', undefined, matched);
}

/**
 * Parse a `read <selector>` slide address into a zero-based slide index, or `undefined` when it
 * isn't an addressable slide reference. Accepts `slide:N` / `slide N` (1-based, human-facing) and a
 * bare 1-based `N`. Conservative: anything else (a name, a range, junk) → `undefined`, so the bridge
 * degrades to `[]` rather than guessing. Pure + exported so the addressing is unit-testable.
 */
export function parseSlideSelector(selector: string): number | undefined {
  const trimmed = selector.trim().toLowerCase();
  const m = /^(?:slide[\s:]*)?(\d{1,4})$/.exec(trimmed);
  if (!m) return undefined;
  const oneBased = Number(m[1]);
  if (!Number.isInteger(oneBased) || oneBased < 1) return undefined;
  return oneBased - 1;
}
