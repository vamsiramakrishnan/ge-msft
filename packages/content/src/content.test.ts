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
    const cjk = '今'.repeat(2000); // ~500 tokens
    const { blocks } = processContent(raw(cjk));
    const budget = 400;
    const chunks = chunkBlocks(blocks, raw(cjk), { maxTokens: budget });
    expect(chunks.length).toBeGreaterThan(1);
    for (const c of chunks) expect(c.meta.tokensEstimate).toBeLessThanOrEqual(budget);
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
