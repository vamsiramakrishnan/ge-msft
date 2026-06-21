import type { ResolvedContext } from '@ge/contracts';
import {
  native,
  toContextNative,
  type Block,
  type NativeContent,
  type ToContextOptions,
} from '@ge/content';

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
