import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import {
  shapesToSlideText,
  slideElementsToBlocks,
  slideElementsToDocStateBlocks,
  slidesToContext,
  selectedSlideToContext,
  searchSlides,
  parseSlideSelector,
  MAX_SEARCH_SLIDES,
  type SlideElement,
} from './capture.js';

describe('powerpoint capture (pure)', () => {
  it('splits shape texts into a title and body lines', () => {
    const { title, body } = shapesToSlideText([
      'SLA & Availability',
      'Contracted: 99.5% monthly\nPolicy standard: 99.9% — gap flagged',
      '   ',
    ]);
    expect(title).toBe('SLA & Availability');
    expect(body).toEqual(['Contracted: 99.5% monthly', 'Policy standard: 99.9% — gap flagged']);
  });

  it('returns an empty title when no shapes carry text', () => {
    expect(shapesToSlideText(['', '   '])).toEqual({ title: '', body: [] });
  });

  it('builds native slide blocks anchored to the slide id, folding notes into the body', () => {
    const slides: SlideElement[] = [
      {
        index: 0,
        slideId: 's9',
        title: 'Risk Summary',
        body: ['Vendor exposure elevated'],
        notes: 'Lead with the gap',
      },
    ];
    const blocks = slideElementsToBlocks(slides);
    expect(blocks[0]?.locator).toBe('slide:s9');
    expect(blocks.some((b) => b.text.includes('Risk Summary'))).toBe(true);
    expect(blocks.some((b) => b.text.includes('Notes: Lead with the gap'))).toBe(true);
  });

  it('falls back to the slide index locator when no slide id is present', () => {
    const blocks = slideElementsToBlocks([{ index: 3, title: 'T', body: ['b'] }]);
    expect(blocks[0]?.locator).toBe('slide:3');
  });

  it('produces valid, anchored context from captured slides', () => {
    const ctx = slidesToContext('pp:deck', 'Whole deck', [
      { index: 0, slideId: 's1', title: 'Vendor QBR', body: ['Quarterly business review'] },
      { index: 1, slideId: 's2', title: 'SLA', body: ['99.5% contracted'] },
    ]);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s1')).toBe(true);
  });

  it('returns nothing for an empty slide set', () => {
    expect(slidesToContext('pp:deck', undefined, [])).toHaveLength(0);
  });

  it('selectedSlideToContext anchors a single slide', () => {
    const ctx = selectedSlideToContext({ index: 4, slideId: 'sx', title: 'One', body: ['a'] });
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:sx')).toBe(true);
  });

  it('slideElementsToDocStateBlocks mirrors the native slide blocks (outline source)', () => {
    const blocks = slideElementsToDocStateBlocks([
      { index: 0, slideId: 's1', title: 'Agenda', body: ['Intro'] },
    ]);
    expect(blocks[0]?.locator).toBe('slide:s1');
    expect(blocks.some((b) => b.text.includes('Agenda'))).toBe(true);
  });
});

describe('powerpoint search (pure)', () => {
  const deck: SlideElement[] = [
    { index: 0, slideId: 's1', title: 'SLA terms', body: ['99.5% contracted'] },
    { index: 1, slideId: 's2', title: 'Roster', body: ['Pat, Sam'] },
    { index: 2, slideId: 's3', title: 'Risk', body: ['SLA gap flagged'] },
  ];

  it('returns slides matching the query (case-insensitive, over title + body)', () => {
    const ctx = searchSlides(deck, 'sla');
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s1')).toBe(true);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s2')).toBe(false);
  });

  it('empty query / no match → []', () => {
    expect(searchSlides(deck, '   ')).toHaveLength(0);
    expect(searchSlides(deck, 'nonexistent-token')).toHaveLength(0);
  });

  it('bounds the result to MAX_SEARCH_SLIDES matches', () => {
    const many: SlideElement[] = Array.from({ length: MAX_SEARCH_SLIDES + 5 }, (_, i) => ({
      index: i,
      slideId: `m${i}`,
      title: 'common token',
      body: [],
    }));
    const ctx = searchSlides(many, 'common');
    const anchors = new Set(ctx.map((c) => c.ref.anchor?.locator));
    expect(anchors.size).toBeLessThanOrEqual(MAX_SEARCH_SLIDES);
  });
});

describe('powerpoint slide selector (pure)', () => {
  it('parses slide:N / slide N / bare N to a zero-based index', () => {
    expect(parseSlideSelector('slide:3')).toBe(2);
    expect(parseSlideSelector('slide 1')).toBe(0);
    expect(parseSlideSelector('5')).toBe(4);
    expect(parseSlideSelector('  SLIDE:2 ')).toBe(1);
  });

  it('rejects unaddressable / out-of-range selectors with undefined', () => {
    expect(parseSlideSelector('')).toBeUndefined();
    expect(parseSlideSelector('Agenda')).toBeUndefined();
    expect(parseSlideSelector('slide:0')).toBeUndefined(); // 1-based; 0 is invalid
    expect(parseSlideSelector('A1:B3')).toBeUndefined();
  });
});
