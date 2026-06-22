import type { ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';
import type { WordParagraph } from './host-port.js';

/** A re-resolved search hit at the host boundary: the matched text + a short surrounding hint. */
export interface WordSearchHit {
  readonly text: string;
  readonly contextHint?: string;
}

/**
 * Pure mapping from Word's native object model into grounding-ready context — no Office.js
 * here, so it's unit-testable. The `WordBridge` reads paragraphs/tables/content-controls via
 * `Word.run` and hands the extracted elements to these functions; they go straight through
 * `@ge/content` (native path, no Markdown round-trip) and carry a `cc:<id>` write-back locator.
 */
export interface WordElement {
  kind: 'heading' | 'paragraph' | 'table';
  text: string;
  level?: number; // heading level, derived from the built-in style
  contentControlId?: number;
  columns?: string[];
  rows?: string[][];
}

export function wordElementsToBlocks(elements: WordElement[]): Block[] {
  const blocks: Block[] = [];
  for (const el of elements) {
    const locator = el.contentControlId !== undefined ? `cc:${el.contentControlId}` : undefined;
    if (el.kind === 'heading') {
      blocks.push(native.heading(el.text, el.level ?? 2, locator));
    } else if (el.kind === 'table' && el.columns && el.rows) {
      blocks.push(native.table({ columns: el.columns, rows: el.rows }, locator));
    } else {
      blocks.push(native.paragraph(el.text, locator));
    }
  }
  return blocks;
}

export function wordDocumentToContext(
  sourceId: string,
  title: string | undefined,
  elements: WordElement[],
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const content: NativeContent = {
    sourceId,
    surface: 'word',
    ...(title ? { title } : {}),
    blocks: wordElementsToBlocks(elements),
  };
  return toContextNative(content, opts);
}

/** A live selection is a single text part (re-resolved at send-time). */
export function wordSelectionToContext(text: string): ResolvedContext[] {
  if (!text.trim()) return [];
  return [
    {
      ref: {
        id: 'word:selection',
        kind: 'selection',
        surface: 'word',
        title: 'Selection',
        preview: text.slice(0, 120),
        live: true,
      },
      value: { as: 'text', text, mimeType: 'text/markdown' },
    },
  ];
}

/** Map a Word built-in style name to a heading level (0 = not a heading). */
export function headingLevel(styleBuiltIn: string): number {
  const m = /Heading\s*(\d)/i.exec(styleBuiltIn);
  return m ? Number(m[1]) : 0;
}

/**
 * Map read-back body paragraphs (text + built-in style) to {@link WordElement}s — headings carry
 * their derived level — exactly as `WordBridge.resolveContext` does. Shared so the `<doc_state>`
 * snapshot's blocks come from the same native mapping as grounding context (ADR-0003).
 */
export function paragraphsToElements(paras: readonly WordParagraph[]): WordElement[] {
  return paras.map((p) => {
    const level = headingLevel(p.styleBuiltIn);
    return level > 0
      ? { kind: 'heading' as const, text: p.text, level }
      : { kind: 'paragraph' as const, text: p.text };
  });
}

/** Body paragraphs → `Block[]` for `buildDocStateSnapshot` (headings get levels + locators). */
export function paragraphsToBlocks(paras: readonly WordParagraph[]): Block[] {
  return wordElementsToBlocks(paragraphsToElements(paras));
}

/**
 * Map bounded, re-resolved `body.search` hits to content-anchored {@link ResolvedContext} — one
 * live text part per hit, anchored by the matched text (the contextHint is folded into the part
 * text as a surrounding cue). Re-resolution happens at the host; this is the pure shaping step.
 * Empty input → `[]`.
 */
export function searchHitsToContext(
  query: string,
  hits: readonly WordSearchHit[],
): ResolvedContext[] {
  const out: ResolvedContext[] = [];
  for (const hit of hits) {
    const text = hit.text.trim();
    if (!text) continue;
    const body =
      hit.contextHint && hit.contextHint.trim() ? `${text}\n\n…${hit.contextHint.trim()}` : text;
    out.push({
      ref: {
        id: 'word:search',
        kind: 'selection',
        surface: 'word',
        title: `Match: ${query}`.slice(0, 80),
        preview: text.slice(0, 120),
        live: true,
        anchor: { matchText: text.slice(0, 120) },
      },
      value: { as: 'text', text: body, mimeType: 'text/markdown' },
    });
  }
  return out;
}
