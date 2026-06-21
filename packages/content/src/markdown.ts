import type { Block, BlockKind } from './model.js';

const HEADING = /^(#{1,6})\s+(.*)$/;
const FENCE = /^(```|~~~)/;
const TABLE_ROW = /^\s*\|.*\|\s*$/;
const QUOTE = /^\s*>/;
const LIST = /^\s*([-*+]|\d+[.)])\s+/;

/**
 * Line-based block parser for Markdown that preserves **character offsets** into the
 * source (so chunks can produce content anchors for write-back). Recognizes ATX
 * headings, fenced code, GitHub-flavored tables, blockquotes, lists, and paragraphs.
 * Intentionally small — markitdown/Docling do the rich format→Markdown conversion
 * upstream (in the bridges); this consumes their output.
 */
export function parseMarkdownBlocks(md: string): Block[] {
  const blocks: Block[] = [];
  const lines = md.split('\n');
  // Precompute the start offset of each line.
  const offsets: number[] = [];
  let cursor = 0;
  for (const line of lines) {
    offsets.push(cursor);
    cursor += line.length + 1; // +1 for the '\n'
  }
  const lineEnd = (i: number): number => offsets[i]! + lines[i]!.length;

  let i = 0;
  const push = (kind: BlockKind, from: number, to: number, level?: number): void => {
    const text = md.slice(offsets[from]!, lineEnd(to)).replace(/\s+$/, '');
    if (text.trim().length === 0) return;
    blocks.push({
      kind,
      text,
      start: offsets[from]!,
      end: lineEnd(to),
      ...(level ? { level } : {}),
    });
  };

  while (i < lines.length) {
    const line = lines[i]!;
    if (line.trim() === '') {
      i++;
      continue;
    }
    // Fenced code: consume to the closing fence.
    if (FENCE.test(line)) {
      const start = i;
      i++;
      while (i < lines.length && !FENCE.test(lines[i]!)) i++;
      push('code', start, Math.min(i, lines.length - 1));
      i++;
      continue;
    }
    const heading = HEADING.exec(line);
    if (heading) {
      push('heading', i, i, heading[1]!.length);
      i++;
      continue;
    }
    if (TABLE_ROW.test(line)) {
      const start = i;
      while (i < lines.length && TABLE_ROW.test(lines[i]!)) i++;
      push('table', start, i - 1);
      continue;
    }
    if (QUOTE.test(line)) {
      const start = i;
      while (i < lines.length && QUOTE.test(lines[i]!)) i++;
      push('quote', start, i - 1);
      continue;
    }
    if (LIST.test(line)) {
      const start = i;
      while (i < lines.length && (LIST.test(lines[i]!) || /^\s+\S/.test(lines[i]!))) i++;
      push('list', start, i - 1);
      continue;
    }
    // Paragraph: run of non-blank lines until a blank or a structural line.
    const start = i;
    while (
      i < lines.length &&
      lines[i]!.trim() !== '' &&
      !HEADING.test(lines[i]!) &&
      !FENCE.test(lines[i]!) &&
      !TABLE_ROW.test(lines[i]!) &&
      !QUOTE.test(lines[i]!) &&
      !LIST.test(lines[i]!)
    ) {
      i++;
    }
    push('paragraph', start, i - 1);
  }
  return blocks;
}

/** Render a structured table (e.g. an Excel range) as a GitHub-flavored Markdown table. */
export function tableToMarkdown(columns: string[], rows: (string | number)[][]): string {
  const head = `| ${columns.join(' | ')} |`;
  const sep = `| ${columns.map(() => '---').join(' | ')} |`;
  const body = rows.map((r) => `| ${r.map((c) => String(c)).join(' | ')} |`).join('\n');
  return [head, sep, body].join('\n');
}
