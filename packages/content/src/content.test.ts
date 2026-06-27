import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import { estimateTokens } from './tokens.js';
import { parseMarkdownBlocks, tableToMarkdown } from './markdown.js';
import { htmlToMarkdown } from './normalize.js';
import { chunkBlocks, splitText } from './chunk.js';
import { contextualizeChunk } from './contextualize.js';
import { processContent, toContext } from './process.js';
import type { RawContent } from './model.js';

const DOC = `# 5. Service Levels

Intro paragraph about the agreement between the parties.

## 5.1 Availability

The services are available 99.5% of the time, measured monthly.

## 5.2 Support

| Tier | Response |
| --- | --- |
| P1 | 1 hour |
| P2 | 4 hours |
`;

function raw(text: string, over: Partial<RawContent> = {}): RawContent {
  return {
    sourceId: 'word:body',
    text,
    format: 'markdown',
    title: 'MSA',
    surface: 'word',
    ...over,
  };
}

describe('tokens', () => {
  it('estimates and floors at word count', () => {
    expect(estimateTokens('')).toBe(0);
    expect(estimateTokens('one two three')).toBeGreaterThanOrEqual(3);
  });

  it('is script-aware: CJK counts ~1 token/char, far denser than Latin', () => {
    // 40 Han chars (no spaces) => ~40 tokens (1/char), vs 40 Latin chars => ~10 (4/char).
    const cjk = '今'.repeat(40);
    const latin = 'a'.repeat(40);
    expect(estimateTokens(cjk)).toBe(40);
    expect(estimateTokens(latin)).toBeLessThanOrEqual(10);
    expect(estimateTokens(cjk)).toBeGreaterThan(estimateTokens(latin));
  });

  it('sums mixed-script populations rather than averaging', () => {
    // 4 Latin chars (~1 token) + 4 Han chars (~4 tokens) => ~5, not ~2.
    expect(estimateTokens('test今日明日')).toBeGreaterThanOrEqual(5);
  });

  it('is deterministic', () => {
    const s = 'The quick brown fox。今日は良い天気です。';
    expect(estimateTokens(s)).toBe(estimateTokens(s));
  });
});

describe('markdown parsing', () => {
  it('splits into structural blocks with offsets', () => {
    const blocks = parseMarkdownBlocks(DOC);
    const kinds = blocks.map((b) => b.kind);
    expect(kinds).toContain('heading');
    expect(kinds).toContain('paragraph');
    expect(kinds).toContain('table');
    // offsets map back to the source text
    const h = blocks.find((b) => b.kind === 'heading')!;
    expect(DOC.slice(h.start, h.end)).toContain('5. Service Levels');
  });

  it('renders a structured table as GFM', () => {
    const md = tableToMarkdown(['Vendor', 'Risk'], [['Acme', 'High']]);
    expect(md).toContain('| Vendor | Risk |');
    expect(md).toContain('| Acme | High |');
  });

  it('reduces basic HTML to markdown', () => {
    expect(htmlToMarkdown('<h2>Title</h2><p>Body <strong>bold</strong></p>')).toContain('## Title');
    expect(htmlToMarkdown('<ul><li>a</li><li>b</li></ul>')).toContain('- a');
  });
});

describe('chunking', () => {
  it('groups blocks under their section breadcrumb', () => {
    const { blocks } = processContent(raw(DOC));
    const chunks = chunkBlocks(blocks, raw(DOC), { maxTokens: 60, sectionBreakLevel: 2 });
    expect(chunks.length).toBeGreaterThan(1);
    const avail = chunks.find((c) => c.text.includes('99.5%'))!;
    expect(avail.meta.sectionPath).toEqual(['5. Service Levels', '5.1 Availability']);
    // every chunk carries a write-back anchor with a locator
    expect(avail.meta.anchor.matchText.length).toBeGreaterThan(0);
    expect(avail.meta.anchor.locator).toMatch(/^chars:\d+-\d+$/);
    expect(avail.meta.anchor.contextHint).toContain('Availability');
  });

  it('keeps a table inside one chunk', () => {
    const { blocks } = processContent(raw(DOC));
    const chunks = chunkBlocks(blocks, raw(DOC), { maxTokens: 40 });
    const tableChunks = chunks.filter((c) => c.text.includes('| P1 |'));
    expect(tableChunks).toHaveLength(1);
  });

  it('recursively splits an oversized paragraph with overlap', () => {
    const long = Array.from({ length: 60 }, (_, i) => `Sentence number ${i} about risk.`).join(' ');
    const { chunks } = processContent(raw(`# Big\n\n${long}`), {
      maxTokens: 50,
      overlapTokens: 10,
    });
    expect(chunks.length).toBeGreaterThan(2);
    for (const c of chunks) expect(c.meta.tokensEstimate).toBeLessThanOrEqual(80);
  });

  // Bug B regression: an oversized sentence is split via wordWrap, whose offsets must carry
  // baseOffset (matching the normal flush path) so char anchors aren't off by block.start.
  it('carries baseOffset through wordWrap for an oversized sentence', () => {
    const text = 'alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima';
    const base = 1000;
    const pieces = splitText(text, 3, 1, base);
    expect(pieces.length).toBeGreaterThan(1);
    // No piece should start before baseOffset — that would mean baseOffset was dropped.
    for (const p of pieces) {
      expect(p.start).toBeGreaterThanOrEqual(base);
      expect(p.end).toBeGreaterThanOrEqual(p.start);
    }
    expect(pieces[0]!.start).toBe(base);
  });

  // Bug C regression: whitespace-free oversized text (e.g. CJK with no spaces) must be
  // hard-split so every chunk respects the token budget.
  it('sub-splits a whitespace-free CJK block to stay within budget', () => {
    const cjk = '今'.repeat(2000); // ~2000 tokens (script-aware: ~1/char)
    const { blocks } = processContent(raw(cjk));
    const budget = 400;
    const chunks = chunkBlocks(blocks, raw(cjk), { maxTokens: budget });
    expect(chunks.length).toBeGreaterThan(1);
    for (const c of chunks) expect(c.meta.tokensEstimate).toBeLessThanOrEqual(budget);
  });

  // Unicode sentence segmentation: split a paragraph into sentences with correct offsets,
  // even where a regex `.!?` splitter would mis-handle abbreviations / non-Latin punctuation.
  it('segments sentences with correct absolute offsets', () => {
    const text = 'First sentence here. Second one follows! And a third? Yes.';
    const base = 500;
    // Tiny budget forces one unit per sentence so we can inspect boundaries.
    const pieces = splitText(text, 2, 0, base);
    expect(pieces.length).toBeGreaterThan(2);
    for (const p of pieces) {
      expect(p.start).toBeGreaterThanOrEqual(base);
      expect(p.end).toBeGreaterThanOrEqual(p.start);
      // The piece text must be recoverable from the source at its (offset - base).
      const slice = text.slice(p.start - base, p.end - base);
      expect(slice).toContain(p.text);
    }
    expect(pieces[0]!.start).toBe(base);
  });

  // CJK has no spaces: a word-granularity segmenter (not /(\s+)/) must still wrap it within
  // budget, falling back to grapheme splitting for the boundary-free run.
  it('wraps a space-free CJK paragraph within budget', () => {
    // Mixed Han sentence repeated, no ASCII spaces anywhere.
    const cjk = '機械学習は自然言語処理を変革する。'.repeat(80);
    const { blocks } = processContent(raw(cjk));
    const budget = 50;
    const chunks = chunkBlocks(blocks, raw(cjk), { maxTokens: budget });
    expect(chunks.length).toBeGreaterThan(1);
    // Soft budget: packing whole sentence units can overshoot by at most one unit's worth.
    // The point is the space-free run *does* get wrapped at all (regex /(\s+)/ never would).
    for (const c of chunks) expect(c.meta.tokensEstimate).toBeLessThanOrEqual(budget + 20);
  });

  // A single token bigger than the budget must hard-split on grapheme clusters without ever
  // cutting a surrogate pair (emoji) or a base + combining-mark sequence.
  it('hard-splits a giant single token on graphemes without breaking surrogate pairs', () => {
    // Emoji (surrogate pairs) + a combining-mark grapheme cluster, no whitespace.
    const giant = '👨‍👩‍👧‍👦🎉🚀é́'.repeat(60);
    const base = 7;
    const pieces = splitText(giant, 5, 0, base);
    expect(pieces.length).toBeGreaterThan(1);
    const rejoined = pieces.map((p) => p.text).join('');
    // No lone surrogate may appear at any slice boundary — round-trips losslessly.
    for (const p of pieces) {
      expect(p.text).not.toMatch(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/); // high surrogate w/o low
      expect(p.text).not.toMatch(/(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/); // low surrogate w/o high
      expect(p.start).toBeGreaterThanOrEqual(base);
    }
    // Concatenating the slices reconstructs the original run exactly (offsets contiguous).
    expect(rejoined).toBe(giant);
    expect(pieces[0]!.start).toBe(base);
  });

  it('does not hard-fail on an odd locale (falls back gracefully)', () => {
    const text = 'Alpha bravo charlie. Delta echo foxtrot.';
    expect(() =>
      chunkBlocks(processContent(raw(text)).blocks, raw(text), {
        maxTokens: 3,
        locale: 'zz-Nonsense-XX',
      }),
    ).not.toThrow();
  });
});

describe('contextualization', () => {
  it('prepends a compact source › section header', () => {
    const { chunks } = processContent(raw(DOC), { maxTokens: 60 });
    const c = chunks.find((x) => x.text.includes('99.5%'))!;
    expect(contextualizeChunk(c)).toMatch(/^\[MSA › 5\. Service Levels › 5\.1 Availability\]/);
  });
});

describe('toContext — attach-ready output', () => {
  it('emits one valid contextualized text part per chunk', () => {
    const ctx = toContext(raw(DOC), { maxTokens: 60 });
    expect(ctx.length).toBeGreaterThan(1);
    for (const c of ctx) {
      expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
      expect(c.value).toMatchObject({ as: 'text', mimeType: 'text/markdown' });
      expect(c.ref.anchor).toBeDefined();
    }
  });

  it('prefers an indexed-document reference when asked (reference-over-inline)', () => {
    const ctx = toContext(
      raw(DOC, { indexedDocumentName: 'projects/x/dataStores/sp/documents/msa' }),
      { preferReference: true },
    );
    expect(ctx).toHaveLength(1);
    expect(ctx[0]!.value).toMatchObject({
      as: 'indexed-document',
      documentName: 'projects/x/dataStores/sp/documents/msa',
    });
  });
});
