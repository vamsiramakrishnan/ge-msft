import type { Anchor } from '@ge/contracts';
import type { Block, Chunk, ChunkOptions, RawContent } from './model.js';
import { estimateTokens } from './tokens.js';

const DEFAULTS = { maxTokens: 400, overlapTokens: 40, sectionBreakLevel: 2 } as const;

/**
 * Structure-aware chunker. Groups blocks under their heading breadcrumb, packs them up
 * to a soft token budget, starts a fresh chunk at section boundaries, and never splits a
 * table. A block that alone exceeds the budget is recursively split (paragraph → sentence
 * → word) with sentence overlap — the LangChain RecursiveCharacterTextSplitter idea, but
 * token-aware and offset-preserving so every chunk keeps a write-back anchor.
 */
export function chunkBlocks(blocks: Block[], raw: RawContent, opts: ChunkOptions = {}): Chunk[] {
  const o = { ...DEFAULTS, ...opts };
  const chunks: Chunk[] = [];
  const headingStack: { level: number; title: string }[] = [];
  let pending: Block[] = [];

  const sectionPath = (): string[] => headingStack.map((h) => h.title);

  const flush = (): void => {
    if (pending.length === 0) return;
    const path = sectionPath();
    const text = pending
      .map((b) => b.text)
      .join('\n\n')
      .trim();
    const start = pending[0]!.start;
    const end = pending[pending.length - 1]!.end;
    emit(chunks, raw, text, start, end, path, dedupeKinds(pending), o);
    pending = [];
  };

  const pendingTokens = (): number => estimateTokens(pending.map((b) => b.text).join('\n\n'));

  for (const block of blocks) {
    if (block.kind === 'heading') {
      const level = block.level ?? 6;
      // A section-level heading closes the current chunk and updates the breadcrumb.
      if (level <= o.sectionBreakLevel) flush();
      while (headingStack.length && headingStack[headingStack.length - 1]!.level >= level) {
        headingStack.pop();
      }
      headingStack.push({ level, title: headingText(block.text) });
      pending.push(block); // keep the heading with its section
      continue;
    }

    const blockTokens = estimateTokens(block.text);
    if (blockTokens > o.maxTokens) {
      // Oversized block: flush what we have, then split this block on its own.
      flush();
      const path = sectionPath();
      for (const piece of splitText(block.text, o.maxTokens, o.overlapTokens, block.start)) {
        emit(chunks, raw, piece.text, piece.start, piece.end, path, [block.kind], o);
      }
      continue;
    }

    if (pendingTokens() + blockTokens > o.maxTokens) flush();
    pending.push(block);
  }
  flush();
  return chunks;
}

function emit(
  chunks: Chunk[],
  raw: RawContent,
  text: string,
  start: number,
  end: number,
  sectionPath: string[],
  kinds: Block['kind'][],
  _o: ChunkOptions,
): void {
  if (text.trim().length === 0) return;
  const index = chunks.length;
  const anchor: Anchor = {
    matchText: firstLine(text).slice(0, 90),
    ...(sectionPath.length ? { contextHint: sectionPath.join(' › ') } : {}),
    locator: `chars:${start}-${end}`,
  };
  chunks.push({
    id: `${raw.sourceId}#${index}`,
    index,
    text,
    meta: {
      sourceId: raw.sourceId,
      ...(raw.title ? { sourceTitle: raw.title } : {}),
      ...(raw.surface ? { surface: raw.surface } : {}),
      sectionPath,
      kinds,
      charStart: start,
      charEnd: end,
      tokensEstimate: estimateTokens(text),
      anchor,
    },
  });
}

/**
 * Recursively split an oversized text by the most semantic separator that helps:
 * paragraphs → sentences → words. Packs pieces to the budget with token overlap and
 * tracks char offsets relative to `baseOffset` so anchors stay valid.
 */
export function splitText(
  text: string,
  maxTokens: number,
  overlapTokens: number,
  baseOffset: number,
): { text: string; start: number; end: number }[] {
  const units = segment(text);
  const out: { text: string; start: number; end: number }[] = [];
  let buf: typeof units = [];
  const bufTokens = (): number => estimateTokens(buf.map((u) => u.text).join(''));

  const flush = (): void => {
    if (buf.length === 0) return;
    const start = baseOffset + buf[0]!.start;
    const end = baseOffset + buf[buf.length - 1]!.end;
    out.push({
      text: buf
        .map((u) => u.text)
        .join('')
        .trim(),
      start,
      end,
    });
    // Carry trailing units as overlap for the next piece.
    const overlap: typeof units = [];
    let acc = 0;
    for (let k = buf.length - 1; k >= 0; k--) {
      acc += estimateTokens(buf[k]!.text);
      overlap.unshift(buf[k]!);
      if (acc >= overlapTokens) break;
    }
    buf = overlap;
  };

  for (const u of units) {
    if (estimateTokens(u.text) > maxTokens) {
      // A single unit still too big (e.g. no sentence breaks): hard-wrap by words.
      flush();
      buf = [];
      for (const w of wordWrap(u, maxTokens)) out.push(w);
      continue;
    }
    if (bufTokens() + estimateTokens(u.text) > maxTokens && buf.length > 0) flush();
    buf.push(u);
  }
  if (buf.length) {
    const start = baseOffset + buf[0]!.start;
    const end = baseOffset + buf[buf.length - 1]!.end;
    out.push({
      text: buf
        .map((u) => u.text)
        .join('')
        .trim(),
      start,
      end,
    });
  }
  return out.filter((p) => p.text.length > 0);
}

interface Unit {
  text: string;
  start: number;
  end: number;
}

/** Split into paragraph-or-sentence units, preserving offsets and trailing whitespace. */
function segment(text: string): Unit[] {
  const units: Unit[] = [];
  // Paragraph boundaries first.
  const paraRe = /[^\n]*(?:\n(?!\n)[^\n]*)*(?:\n\n+|$)/g;
  let m: RegExpExecArray | null;
  while ((m = paraRe.exec(text)) !== null) {
    if (m[0].length === 0) {
      paraRe.lastIndex++;
      continue;
    }
    const paraStart = m.index;
    // Sentence boundaries within the paragraph.
    const sentRe = /[^.!?]*[.!?]+[\s)"']*|\S[^.!?]*$/g;
    let s: RegExpExecArray | null;
    let any = false;
    while ((s = sentRe.exec(m[0])) !== null) {
      if (s[0].trim().length === 0) continue;
      any = true;
      units.push({
        text: s[0],
        start: paraStart + s.index,
        end: paraStart + s.index + s[0].length,
      });
    }
    if (!any) units.push({ text: m[0], start: paraStart, end: paraStart + m[0].length });
  }
  return units;
}

function wordWrap(u: Unit, maxTokens: number): { text: string; start: number; end: number }[] {
  const out: { text: string; start: number; end: number }[] = [];
  const words = u.text.split(/(\s+)/);
  let buf = '';
  let start = u.start;
  let cursor = u.start;
  for (const w of words) {
    if (estimateTokens(buf + w) > maxTokens && buf.trim().length > 0) {
      out.push({ text: buf.trim(), start, end: cursor });
      buf = '';
      start = cursor;
    }
    buf += w;
    cursor += w.length;
  }
  if (buf.trim().length > 0) out.push({ text: buf.trim(), start, end: cursor });
  return out;
}

function headingText(raw: string): string {
  return raw.replace(/^#{1,6}\s+/, '').trim();
}

function firstLine(text: string): string {
  return text
    .replace(/^#{1,6}\s+/, '')
    .split('\n')[0]!
    .trim();
}

function dedupeKinds(blocks: Block[]): Block['kind'][] {
  return [...new Set(blocks.map((b) => b.kind))];
}
