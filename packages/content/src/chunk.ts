import type { Anchor } from '@ge/contracts';
import type { Block, Chunk, ChunkOptions, SourceMeta } from './model.js';
import { estimateTokens } from './tokens.js';

const DEFAULTS = { maxTokens: 400, overlapTokens: 40, sectionBreakLevel: 2 } as const;

/**
 * Structure-aware chunker. Groups blocks under their heading breadcrumb, packs them up
 * to a soft token budget, starts a fresh chunk at section boundaries, and never splits a
 * table. A block that alone exceeds the budget is recursively split (paragraph → sentence
 * → word) with sentence overlap. Every chunk keeps a write-back anchor: a native host
 * `locator` when the source provided one, else `chars:start-end` from the string path.
 *
 * Works the same whether `blocks` came from the Office object model (native) or the
 * Markdown parser (string fallback).
 */
export function chunkBlocks(blocks: Block[], meta: SourceMeta, opts: ChunkOptions = {}): Chunk[] {
  const o = { ...DEFAULTS, ...opts };
  const chunks: Chunk[] = [];
  const headingStack: { level: number; title: string }[] = [];
  let pending: Block[] = [];

  const sectionPath = (): string[] => headingStack.map((h) => h.title);

  const flush = (): void => {
    if (pending.length === 0) return;
    emit(chunks, meta, pending, sectionPath(), o);
    pending = [];
  };

  const pendingTokens = (): number => estimateTokens(pending.map((b) => b.text).join('\n\n'));

  for (const block of blocks) {
    if (block.kind === 'heading') {
      const level = block.level ?? 6;
      if (level <= o.sectionBreakLevel) flush();
      while (headingStack.length && headingStack[headingStack.length - 1]!.level >= level) {
        headingStack.pop();
      }
      headingStack.push({ level, title: headingText(block.text) });
      pending.push(block);
      continue;
    }

    const blockTokens = estimateTokens(block.text);
    // Tables and other native structured blocks are kept whole even if large.
    if (blockTokens > o.maxTokens && block.kind !== 'table') {
      flush();
      const path = sectionPath();
      const base = block.start ?? 0;
      for (const piece of splitText(block.text, o.maxTokens, o.overlapTokens, base)) {
        const part: Block = {
          kind: block.kind,
          text: piece.text,
          start: block.start === undefined ? undefined : piece.start,
          end: block.start === undefined ? undefined : piece.end,
          ...(block.locator ? { locator: block.locator } : {}),
        };
        emit(chunks, meta, [part], path, o);
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
  meta: SourceMeta,
  blocks: Block[],
  sectionPath: string[],
  _o: ChunkOptions,
): void {
  const text = blocks
    .map((b) => b.text)
    .join('\n\n')
    .trim();
  if (text.length === 0) return;
  const index = chunks.length;
  const first = blocks[0]!;
  const last = blocks[blocks.length - 1]!;
  const locator = locatorFor(blocks);
  const anchor: Anchor = {
    matchText: firstLine(text).slice(0, 90),
    ...(sectionPath.length ? { contextHint: sectionPath.join(' › ') } : {}),
    ...(locator ? { locator } : {}),
  };
  const hasOffsets = first.start !== undefined && last.end !== undefined;
  chunks.push({
    id: `${meta.sourceId}#${index}`,
    index,
    text,
    meta: {
      sourceId: meta.sourceId,
      ...(meta.title ? { sourceTitle: meta.title } : {}),
      ...(meta.surface ? { surface: meta.surface } : {}),
      sectionPath,
      kinds: [...new Set(blocks.map((b) => b.kind))],
      ...(hasOffsets ? { charStart: first.start, charEnd: last.end } : {}),
      tokensEstimate: estimateTokens(text),
      anchor,
    },
  });
}

/** Prefer a native locator from the chunk's blocks; else fall back to a char range. */
function locatorFor(blocks: Block[]): string | undefined {
  const native = blocks.find((b) => b.locator)?.locator;
  if (native) return native;
  const first = blocks[0]!;
  const last = blocks[blocks.length - 1]!;
  if (first.start !== undefined && last.end !== undefined)
    return `chars:${first.start}-${last.end}`;
  return undefined;
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
    out.push({
      text: buf
        .map((u) => u.text)
        .join('')
        .trim(),
      start: baseOffset + buf[0]!.start,
      end: baseOffset + buf[buf.length - 1]!.end,
    });
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
      flush();
      buf = [];
      for (const w of wordWrap(u, maxTokens))
        out.push({ text: w.text, start: baseOffset + w.start, end: baseOffset + w.end });
      continue;
    }
    if (bufTokens() + estimateTokens(u.text) > maxTokens && buf.length > 0) flush();
    buf.push(u);
  }
  if (buf.length) {
    out.push({
      text: buf
        .map((u) => u.text)
        .join('')
        .trim(),
      start: baseOffset + buf[0]!.start,
      end: baseOffset + buf[buf.length - 1]!.end,
    });
  }
  return out.filter((p) => p.text.length > 0);
}

interface Unit {
  text: string;
  start: number;
  end: number;
}

function segment(text: string): Unit[] {
  const units: Unit[] = [];
  const paraRe = /[^\n]*(?:\n(?!\n)[^\n]*)*(?:\n\n+|$)/g;
  let m: RegExpExecArray | null;
  while ((m = paraRe.exec(text)) !== null) {
    if (m[0].length === 0) {
      paraRe.lastIndex++;
      continue;
    }
    const paraStart = m.index;
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
  const push = (): void => {
    if (buf.trim().length === 0) return;
    out.push({ text: buf.trim(), start, end: cursor });
    buf = '';
    start = cursor;
  };
  for (const w of words) {
    // A single "word" that alone exceeds the budget (long URL/base64, or whitespace-free
    // CJK text) won't be broken by buffer-flush alone — hard-split it by character so each
    // slice stays within budget. Carry char offsets through so anchors stay valid.
    if (estimateTokens(w) > maxTokens) {
      push();
      for (const slice of hardSplit(w, cursor, maxTokens)) out.push(slice);
      cursor += w.length;
      start = cursor;
      continue;
    }
    if (estimateTokens(buf + w) > maxTokens && buf.trim().length > 0) push();
    buf += w;
    cursor += w.length;
  }
  push();
  return out;
}

/**
 * Split a whitespace-free run into budget-sized character slices. Estimates the per-slice
 * char count from the token budget and the run's own token density, then bisects down until
 * each slice is within budget. Offsets are absolute (anchored at `offset`).
 */
function hardSplit(
  word: string,
  offset: number,
  maxTokens: number,
): { text: string; start: number; end: number }[] {
  const out: { text: string; start: number; end: number }[] = [];
  const tokens = estimateTokens(word);
  if (tokens <= maxTokens || word.length <= 1) {
    out.push({ text: word, start: offset, end: offset + word.length });
    return out;
  }
  // Conservative initial slice length, then shrink any slice still over budget.
  let sliceLen = Math.max(1, Math.floor((word.length * maxTokens) / tokens));
  let pos = 0;
  while (pos < word.length) {
    let len = Math.min(sliceLen, word.length - pos);
    while (len > 1 && estimateTokens(word.slice(pos, pos + len)) > maxTokens) {
      len = Math.floor(len / 2);
    }
    if (len < 1) len = 1;
    out.push({ text: word.slice(pos, pos + len), start: offset + pos, end: offset + pos + len });
    sliceLen = len;
    pos += len;
  }
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
