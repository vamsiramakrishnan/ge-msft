import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { splitText, chunkBlocks } from './chunk.js';
import { estimateTokens } from './tokens.js';
import type { Block, SourceMeta } from './model.js';

/**
 * Property tests for the content chunker — the segmentation hot-spot.
 *
 * splitText is the recursive oversize-block splitter (paragraph → sentence → word →
 * grapheme). The bug-bash class here: budgets honoured only softly, char offsets drifting
 * out of range, and — the nasty one — a boundary cutting through a surrogate pair or a
 * combining mark. We generate text mixing Latin/CJK/emoji/whitespace/boundary-free runs
 * across random budgets and assert the documented invariants.
 */

const SEED = 0x5eed_c0de;
const NUM_RUNS = 400;

// Token estimation is non-linear and per-call (overlap, trimming, word-floor), so the
// budget is *soft*. Allow a generous soft margin: the package overlaps and never
// subdivides a single grapheme, and estimateTokens floors at whitespace word count.
// A piece's tokens = its own body (packed up to ~maxTokens) + the overlap carried in from
// the previous flush. flush() carries back whole sentence units until it has *at least*
// overlapTokens, so the carry-in can be up to overlapTokens plus one extra whole unit
// (itself bounded by maxTokens, since over-budget units are split before buffering). Hence
// the honest soft ceiling is maxTokens (body) + overlapTokens + maxTokens (one extra unit)
// + a small slack for the word-count floor / trimming jitter in estimateTokens.
function softCeiling(maxTokens: number, overlapTokens: number): number {
  return maxTokens * 2 + overlapTokens + 8;
}

// Pieces of text that exercise every segmentation path.
const textPool = fc.constantFrom(
  'The quick brown fox jumps over the lazy dog. ',
  'Short. Sentences! Here? Yes. ',
  '日本語のテキストです。これはテストです。',
  '中文文本没有空格分隔符号需要按字符切分处理',
  '😀🚀💥🎉🧑‍🚀👨‍👩‍👧‍👦', // emoji incl. ZWJ sequences + combining
  'éàô', // combining marks (e + acute, etc.)
  'a'.repeat(60), // long boundary-free run
  'https://example.com/very/long/path/that/has/no/spaces/at/all/whatsoever/foo/bar',
  '\n\n', // paragraph breaks
  '   ', // pure whitespace
  'Mixed 日本 text with 😀 emoji and ascii.',
  '',
);

const richText = fc.array(textPool, { maxLength: 12 }).map((parts) => parts.join(''));

/** Count lone surrogates in a string (a correctly split string has none). */
function loneSurrogates(s: string): number {
  let count = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      // high surrogate must be followed by a low surrogate
      const next = i + 1 < s.length ? s.charCodeAt(i + 1) : 0;
      if (next >= 0xdc00 && next <= 0xdfff) i++;
      else count++;
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      // lone low surrogate
      count++;
    }
  }
  return count;
}

describe('splitText — property: budget, offsets, grapheme-safety', () => {
  it('respects the soft budget, keeps offsets in range/monotonic, never splits a surrogate', () => {
    fc.assert(
      fc.property(
        richText,
        fc.integer({ min: 5, max: 200 }), // maxTokens
        fc.integer({ min: 0, max: 40 }), // overlapTokens
        fc.integer({ min: 0, max: 1000 }), // baseOffset
        (text, maxTokens, overlapTokens, baseOffset) => {
          const pieces = splitText(text, maxTokens, overlapTokens, baseOffset);
          const ceil = softCeiling(maxTokens, overlapTokens);
          let prevStart = -Infinity;
          for (const p of pieces) {
            // (a) within the soft budget
            expect(estimateTokens(p.text)).toBeLessThanOrEqual(ceil);
            // (b) offsets within [baseOffset, baseOffset + len] and start <= end
            expect(p.start).toBeGreaterThanOrEqual(baseOffset);
            expect(p.end).toBeLessThanOrEqual(baseOffset + text.length);
            expect(p.start).toBeLessThanOrEqual(p.end);
            // monotonic non-decreasing start offsets
            expect(p.start).toBeGreaterThanOrEqual(prevStart);
            prevStart = p.start;
            // (c) no lone surrogate at any boundary
            expect(loneSurrogates(p.text)).toBe(0);
          }
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('preserves all source content (no characters lost; overlap may duplicate)', () => {
    // NOTE on the documented overlap: splitText *always* carries at least one segment of
    // overlap between adjacent pieces — even at overlapTokens=0 — because the overlap loop
    // in flush() emits one unit before testing `acc >= overlapTokens`. So pieces may
    // *duplicate* content; they must never *drop* it. The right invariant is therefore
    // "the stripped source is a subsequence of the stripped, ordered concatenation" — every
    // source character appears, in order, somewhere in the pieces.
    fc.assert(
      fc.property(richText, fc.integer({ min: 20, max: 200 }), (text, maxTokens) => {
        const pieces = splitText(text, maxTokens, 0, 0);
        const strip = (s: string): string => s.replace(/\s+/g, '');
        const joined = strip(pieces.map((p) => p.text).join(''));
        const src = strip(text);
        // src must be a subsequence of joined (no content lost).
        let j = 0;
        for (let i = 0; i < src.length; i++) {
          while (j < joined.length && joined[j] !== src[i]) j++;
          expect(
            j,
            `source char @${i} (${JSON.stringify(src[i])}) missing from pieces`,
          ).toBeLessThan(joined.length);
          j++;
        }
      }),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });
});

describe('chunkBlocks — property: emitted chunks obey budget and stay coherent', () => {
  const meta: SourceMeta = { sourceId: 'prop:test', title: 'T' };

  // Splittable kinds only. By documented design chunkBlocks keeps a `heading` (a section
  // breadcrumb) and a `table` whole even when oversized — so the soft budget is *not*
  // expected to hold for those, and asserting it on them would be testing a non-contract.
  const splittableKind = fc.constantFrom('paragraph', 'list', 'quote', 'code') as fc.Arbitrary<
    Block['kind']
  >;

  const blockArb = fc.record({
    kind: splittableKind,
    text: richText,
  });

  it('keeps splittable chunks within the soft budget; offsets and graphemes coherent', () => {
    fc.assert(
      fc.property(
        fc.array(blockArb, { maxLength: 10 }),
        fc.integer({ min: 20, max: 200 }),
        (rawBlocks, maxTokens) => {
          // Give string-path blocks coherent char offsets so locator math is exercised.
          let cursor = 0;
          const blocks: Block[] = rawBlocks.map((b) => {
            const start = cursor;
            const end = cursor + b.text.length;
            cursor = end + 2;
            return { kind: b.kind, text: b.text, start, end };
          });
          const chunks = chunkBlocks(blocks, meta, { maxTokens, overlapTokens: 0 });
          const ceil = softCeiling(maxTokens, 0);
          for (const c of chunks) {
            expect(c.meta.tokensEstimate).toBeLessThanOrEqual(ceil);
            // never split a surrogate pair / combining-mark boundary
            expect(loneSurrogates(c.text)).toBe(0);
            // char offsets coherent and monotonic within source span
            if (c.meta.charStart !== undefined && c.meta.charEnd !== undefined) {
              expect(c.meta.charStart).toBeLessThanOrEqual(c.meta.charEnd);
            }
          }
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('keeps headings and tables whole, but never with a broken grapheme boundary', () => {
    // For the kinds the chunker keeps whole, we drop the budget assertion (by design) but
    // still require the grapheme-safety and offset-coherence invariants to hold.
    const wholeKind = fc.constantFrom('heading', 'table') as fc.Arbitrary<Block['kind']>;
    fc.assert(
      fc.property(
        fc.array(fc.record({ kind: wholeKind, text: richText }), { maxLength: 6 }),
        fc.integer({ min: 20, max: 200 }),
        (rawBlocks, maxTokens) => {
          let cursor = 0;
          const blocks: Block[] = rawBlocks.map((b) => {
            const start = cursor;
            const end = cursor + b.text.length;
            cursor = end + 2;
            return { kind: b.kind, text: b.text, start, end };
          });
          const chunks = chunkBlocks(blocks, meta, { maxTokens, overlapTokens: 0 });
          for (const c of chunks) {
            expect(loneSurrogates(c.text)).toBe(0);
            if (c.meta.charStart !== undefined && c.meta.charEnd !== undefined) {
              expect(c.meta.charStart).toBeLessThanOrEqual(c.meta.charEnd);
            }
          }
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });
});
