import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import { pageElementToBlocks, pageToContext, type PageElement } from './capture.js';

describe('onenote capture (pure)', () => {
  const page: PageElement = {
    pageId: 'p7',
    title: 'Source review',
    paragraphs: ['Northwind MSA v3 is current.', '   ', 'ISO 27001 valid through Nov 2026.'],
  };

  it('builds title + paragraph blocks anchored to the page id, dropping blanks', () => {
    const blocks = pageElementToBlocks(page);
    expect(blocks[0]?.kind).toBe('heading');
    expect(blocks.every((b) => b.locator === 'page:p7')).toBe(true);
    expect(blocks.filter((b) => b.kind === 'paragraph')).toHaveLength(2);
  });

  it('produces valid, anchored context from a page', () => {
    const ctx = pageToContext(page);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'page:p7')).toBe(true);
  });

  it('returns nothing for an empty page', () => {
    expect(pageToContext({ pageId: 'p0', title: '', paragraphs: ['  '] })).toHaveLength(0);
  });
});
