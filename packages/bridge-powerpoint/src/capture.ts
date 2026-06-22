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
