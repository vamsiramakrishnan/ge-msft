import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import {
  shapesToSlideText,
  slideElementsToBlocks,
  slideElementsToDocStateBlocks,
  slidesToContext,
  selectedSlideToContext,
  searchSlides,
  mayNameSlide,
  parseSlideSelector,
  slideIndexForSelector,
  slideShapeListing,
  MAX_SEARCH_SLIDES,
  tableValuesToText,
  type SlideElement,
} from './capture.js';

describe('powerpoint capture (pure)', () => {
  it('flattens a table grid into one line per row, dropping blank rows', () => {
    expect(
      tableValuesToText([
        ['Site', 'Budget'],
        ['Madrid', ' $310K\r\n(approved) '],
        ['', '  '],
        ['Total', '$830K'],
      ]),
    ).toBe('Site | Budget\nMadrid | $310K (approved)\nTotal | $830K');
    expect(tableValuesToText([])).toBe('');
  });

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

  it('slideElementsToDocStateBlocks lists each slide once, by its title, without its body', () => {
    const blocks = slideElementsToDocStateBlocks([
      { index: 0, slideId: 's1', title: 'Agenda', body: ['Intro', 'Scope'] },
      { index: 1, slideId: 's2', title: '', body: ['Only a body'] },
    ]);
    expect(blocks.map((b) => [b.kind, b.text, b.locator])).toEqual([
      ['heading', '## Agenda', 'slide:s1'],
      ['heading', '## Slide 2', 'slide:s2'],
    ]);
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

describe('slide selectors for read', () => {
  const ids = ['256#', '257#', '300#'];

  it('resolves host ids, pp:slide:/slide: refs, numbers and last', () => {
    expect(slideIndexForSelector(ids, '257#')).toBe(1);
    expect(slideIndexForSelector(ids, 'pp:slide:300#')).toBe(2);
    expect(slideIndexForSelector(ids, 'slide:257#')).toBe(1);
    expect(slideIndexForSelector(ids, 'slide 1')).toBe(0);
    expect(slideIndexForSelector(ids, 'pp:slide:2')).toBe(1);
    expect(slideIndexForSelector(ids, 'last')).toBe(2);
    expect(slideIndexForSelector(ids, '4')).toBeUndefined();
    expect(slideIndexForSelector(ids, 'Agenda')).toBeUndefined();
    expect(slideIndexForSelector([], 'last')).toBeUndefined();
  });

  it('screens out selectors that cannot name a slide before any host call', () => {
    expect(mayNameSlide('Agenda')).toBe(false);
    expect(mayNameSlide('A1:B3')).toBe(false);
    expect(mayNameSlide('slide:1')).toBe(true);
    expect(mayNameSlide('pp:slide:256#')).toBe(true);
    expect(mayNameSlide('256#')).toBe(true);
    expect(mayNameSlide('last')).toBe(true);
  });
});

describe('slideShapeListing', () => {
  it('lists ids, types, the title and clipped single-line text', () => {
    const lines = slideShapeListing({
      index: 0,
      slideId: '256#',
      title: 'FY26 Plan',
      body: [],
      shapes: [
        { shapeId: '2', text: 'FY26 Plan', type: 'Placeholder', isTitle: true },
        { shapeId: '3', text: 'line one\nline two', type: 'TextBox' },
        { shapeId: '4', text: 'x'.repeat(100) },
      ],
    });
    expect(lines[0]).toBe(
      'Shapes on slide 1 (for shape commands use slide=1 and shape=<id>; shape=title is id 2):',
    );
    expect(lines[1]).toBe('- id 2 · Placeholder · title · "FY26 Plan"');
    expect(lines[2]).toBe('- id 3 · TextBox · "line one line two"');
    expect(lines[3]).toBe(`- id 4 · "${'x'.repeat(80)}…"`);
  });

  it('is empty when the slide has no shapes', () => {
    expect(slideShapeListing({ index: 0, title: '', body: [] })).toEqual([]);
  });
});
