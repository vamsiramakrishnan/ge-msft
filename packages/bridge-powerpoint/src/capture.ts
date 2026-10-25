import type { ContextRef, ResolvedContext } from '@ge/contracts';
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
export interface ShapeElement {
  /** Stable host shape id within the slide. */
  shapeId: string;
  /** Text read from the shape's text frame; empty for non-text shapes. */
  text: string;
  /** Host shape type (`Placeholder`, `TextBox`, `GeometricShape`, `Image`, `Table`…), when read. */
  type?: string;
  /** True for the shape `shape=title` resolves to on this slide. */
  isTitle?: boolean;
}

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
  /** Addressable shapes on the slide, when the bridge has read them. */
  shapes?: ShapeElement[];
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

/**
 * A native table's cell grid (`Table.values`) → one text line per row, cells joined by ` | `, so it
 * flows through {@link shapesToSlideText} like any other shape text. Line breaks inside a cell
 * become spaces (a row must stay one line); rows whose cells are all blank are dropped.
 */
export function tableValuesToText(values: string[][]): string {
  return values
    .map((row) => row.map((cell) => (cell ?? '').replace(/\s+/g, ' ').trim()))
    .filter((row) => row.some((cell) => cell.length > 0))
    .map((row) => row.join(' | '))
    .join('\n');
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

/**
 * The shape listing a single-slide `read` appends, so the model can address shape commands
 * (`/format-shape slide=N shape=<id>`) without guessing ids. Shape text is document content: it is
 * quoted, single-line and clipped, like any other grounding text.
 */
export function slideShapeListing(slide: SlideElement): string[] {
  const shapes = slide.shapes ?? [];
  if (shapes.length === 0) return [];
  const title = shapes.find((shape) => shape.isTitle);
  const header =
    `Shapes on slide ${slide.index + 1} (for shape commands use slide=${slide.index + 1} and ` +
    `shape=<id>${title ? `; shape=title is id ${title.shapeId}` : ''}):`;
  return [
    header,
    ...shapes.map((shape) => {
      const text = shape.text.replace(/\s+/g, ' ').trim();
      return [
        `- id ${shape.shapeId}`,
        shape.type,
        shape.isTitle ? 'title' : undefined,
        text ? JSON.stringify(text.length > 80 ? `${text.slice(0, 80)}…` : text) : undefined,
      ]
        .filter(Boolean)
        .join(' · ');
    }),
  ];
}

/** A single selected slide → context (same mapping as `slidesToContext`). */
export function selectedSlideToContext(slide: SlideElement): ResolvedContext[] {
  return slidesToContext(`pp:slide:${slide.slideId ?? slide.index}`, undefined, [slide]);
}

export function slideContextRef(
  slide: Pick<SlideElement, 'index' | 'slideId' | 'title'>,
): ContextRef {
  const slideId = slide.slideId ?? String(slide.index);
  return {
    id: `pp:slide:${slideId}`,
    kind: 'slide',
    surface: 'powerpoint',
    title: `Slide ${slide.index + 1}${slide.title ? `: ${slide.title.slice(0, 60)}` : ''}`,
    ...(slide.title ? { preview: slide.title.slice(0, 120) } : {}),
    live: true,
    anchor: { matchText: slide.title ?? `Slide ${slide.index + 1}`, locator: `slide:${slideId}` },
    hostRef: { type: 'powerpoint.slide', slideId },
  };
}

/**
 * A shape as attachable context. Titled by its host id, never its position: the model copies the
 * number from the title into `pp:shape:<slide>:<shape>`, and "Shape 1" (the first shape, often a
 * decorative line) read as shape id 1 targeted nothing, or the wrong shape.
 */
export function shapeContextRef(
  slide: Pick<SlideElement, 'index' | 'slideId'>,
  shape: ShapeElement,
): ContextRef {
  const slideId = slide.slideId ?? String(slide.index);
  return {
    id: `pp:shape:${slideId}:${shape.shapeId}`,
    kind: 'shape',
    surface: 'powerpoint',
    title: `Shape ${shape.shapeId} on slide ${slide.index + 1}`,
    ...(shape.text.trim() ? { preview: shape.text.trim().slice(0, 120) } : {}),
    live: true,
    anchor: {
      matchText: shape.text.trim() || shape.shapeId,
      locator: `pp:shape:${slideId}:${shape.shapeId}`,
    },
    hostRef: { type: 'powerpoint.shape', slideId, shapeId: shape.shapeId },
  };
}

export function selectedShapeToContext(
  slide: Pick<SlideElement, 'index' | 'slideId'>,
  shape: ShapeElement,
): ResolvedContext[] {
  const text = shape.text.trim();
  if (!text) return [];
  const ref = shapeContextRef(slide, shape);
  return [
    {
      ref,
      value: { as: 'text', text, mimeType: 'text/plain' },
    },
  ];
}

/**
 * Captured slides → the `Block[]` the surface-agnostic `buildDocStateSnapshot` consumes for the
 * `<doc_state>` outline/inventory (ADR-0003 Layer B element 1): one `slide:<id>` title heading per
 * slide, which the builder numbers in order, so `slides` must start at the deck's first slide. Body
 * lines are left to `read slide:N`. Listing each bullet as well filled the 60-entry inventory by
 * about slide 11 (live 2026-10-09, a 15-slide deck): a slide just added at the end was missing, and
 * the model spent its remaining turns searching for it instead of adding the table it was asked for.
 */
export function slideElementsToDocStateBlocks(slides: SlideElement[]): Block[] {
  return slides.flatMap((s, i) => native.slide(i, s.title, [], s.slideId));
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

/**
 * Whether a `read` selector could name a slide — `slide:`/`pp:slide:` refs, `last`, or a token with a
 * digit (a slide number or host id such as `256#`). Checked BEFORE calling the host, so a name like
 * `Agenda` degrades to `[]` without a round-trip.
 */
export function mayNameSlide(selector: string): boolean {
  const key = selector.trim();
  return /^(?:pp:)?slide\b/i.test(key) || /^last$/i.test(key) || /^[\w#-]*\d[\w#-]*$/.test(key);
}

/**
 * Resolve a `read` selector against the deck's slide ids (in order): an exact host id (bare or as
 * `slide:<id>` / `pp:slide:<id>`), `last`, else a 1-based slide number ({@link parseSlideSelector}).
 * Ids win over numbers, matching the write path's slide resolution.
 */
export function slideIndexForSelector(
  ids: readonly string[],
  selector: string,
): number | undefined {
  const key = selector.trim().replace(/^(?:pp:)?slide[\s:]*/i, '');
  const byId = ids.indexOf(key);
  if (byId >= 0) return byId;
  if (/^last$/i.test(key)) return ids.length > 0 ? ids.length - 1 : undefined;
  const index = parseSlideSelector(key);
  return index !== undefined && index < ids.length ? index : undefined;
}
