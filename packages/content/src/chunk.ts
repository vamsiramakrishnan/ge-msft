import type { Anchor } from '@ge/contracts';
import type { Block, Chunk, ChunkOptions, SourceMeta } from './model.js';
import { estimateTokens } from './tokens.js';

const DEFAULTS = { maxTokens: 400, overlapTokens: 40, sectionBreakLevel: 2 } as const;

/** Default locale for Unicode segmentation. Configurable per call so it can be tuned later. */
const DEFAULT_LOCALE = 'en';

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
      for (const piece of splitText(block.text, o.maxTokens, o.overlapTokens, base, o.locale)) {
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
  locale: string = DEFAULT_LOCALE,
): { text: string; start: number; end: number }[] {
  const units = segment(text, locale);
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
      for (const w of wordWrap(u, maxTokens, locale))
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

/** Lazily-built, memoized segmenters keyed by `granularity|locale`. */
const segmenterCache = new Map<string, Intl.Segmenter>();

function getSegmenter(
  granularity: Intl.SegmenterOptions['granularity'],
  locale: string,
): Intl.Segmenter {
  const key = `${granularity}|${locale}`;
  let seg = segmenterCache.get(key);
  if (!seg) {
    // Don't hard-fail on an odd/unknown locale — Intl falls back to a default; if even
    // construction throws (malformed tag), retry with the library default.
    try {
      seg = new Intl.Segmenter(locale, { granularity });
    } catch {
      seg = new Intl.Segmenter(DEFAULT_LOCALE, { granularity });
    }
    segmenterCache.set(key, seg);
  }
  return seg;
}

/**
 * Split text into sentence-granularity units with correct UTF-16 offsets.
 *
 * Paragraph boundaries (`\n\n`) are structural and kept as hard splits so a chunk never
 * straddles them; within each paragraph we use `Intl.Segmenter(locale, { granularity:
 * 'sentence' })` — which segments correctly across scripts (CJK, Thai, etc.) where a regex
 * `.!?`-based splitter fails. Each segment carries its `index` (UTF-16 offset) through as
 * the absolute char offset, preserving the baseOffset discipline downstream.
 */
function segment(text: string, locale: string): Unit[] {
  const units: Unit[] = [];
  const sentenceSeg = getSegmenter('sentence', locale);
  // Walk paragraphs first; `paraRe` matches a run up to (and including) a blank-line break.
  const paraRe = /[^\n]*(?:\n(?!\n)[^\n]*)*(?:\n\n+|$)/g;
  let m: RegExpExecArray | null;
  while ((m = paraRe.exec(text)) !== null) {
    if (m[0].length === 0) {
      paraRe.lastIndex++;
      continue;
    }
    const paraStart = m.index;
    const para = m[0];
    let any = false;
    for (const { segment: s, index } of sentenceSeg.segment(para)) {
      if (s.trim().length === 0) continue;
      any = true;
      units.push({ text: s, start: paraStart + index, end: paraStart + index + s.length });
    }
    if (!any) units.push({ text: para, start: paraStart, end: paraStart + para.length });
  }
  return units;
}

/**
 * Wrap an over-budget sentence at word boundaries using `Intl.Segmenter(locale, {
 * granularity: 'word' })` (replaces a `/(\s+)/` split, which is wrong for space-free
 * scripts). A single "word-like" segment that alone exceeds the budget — a long URL,
 * base64 blob, or a boundary-free CJK run — is hard-split on grapheme clusters so we never
 * cut a surrogate pair or combining mark. Offsets stay absolute (anchored at `u.start`).
 */
function wordWrap(
  u: Unit,
  maxTokens: number,
  locale: string,
): { text: string; start: number; end: number }[] {
  const out: { text: string; start: number; end: number }[] = [];
  const wordSeg = getSegmenter('word', locale);
  let buf = '';
  let start = u.start;
  let cursor = u.start;
  const push = (): void => {
    if (buf.trim().length === 0) {
      // Discard buffered pure-whitespace but still advance `start` past it.
      buf = '';
      start = cursor;
      return;
    }
    out.push({ text: buf.trim(), start, end: cursor });
    buf = '';
    start = cursor;
  };
  for (const { segment: w } of wordSeg.segment(u.text)) {
    if (estimateTokens(w) > maxTokens) {
      push();
      for (const slice of graphemeSplit(w, cursor, maxTokens, locale)) out.push(slice);
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
 * Hard-split a boundary-free run into budget-sized slices on *grapheme cluster* boundaries
 * via `Intl.Segmenter(locale, { granularity: 'grapheme' })`. This guarantees we never split
 * a surrogate pair (e.g. an emoji) or a base + combining-mark sequence. We greedily pack
 * graphemes up to the token budget; a single grapheme that alone exceeds the budget is
 * emitted on its own (we never subdivide one). Offsets are absolute (anchored at `offset`).
 */
function graphemeSplit(
  run: string,
  offset: number,
  maxTokens: number,
  locale: string,
): { text: string; start: number; end: number }[] {
  const out: { text: string; start: number; end: number }[] = [];
  if (estimateTokens(run) <= maxTokens) {
    out.push({ text: run, start: offset, end: offset + run.length });
    return out;
  }
  const graphemeSeg = getSegmenter('grapheme', locale);
  let buf = '';
  let bufStart = offset;
  let cursor = offset;
  const flush = (): void => {
    if (buf.length === 0) return;
    out.push({ text: buf, start: bufStart, end: cursor });
    buf = '';
    bufStart = cursor;
  };
  for (const { segment: g } of graphemeSeg.segment(run)) {
    // Adding this grapheme would overflow — flush what we have first (unless empty).
    if (buf.length > 0 && estimateTokens(buf + g) > maxTokens) flush();
    buf += g;
    cursor += g.length;
  }
  flush();
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
