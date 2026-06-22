import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import {
  shapesToSlideText,
  slideElementsToBlocks,
  slidesToContext,
  selectedSlideToContext,
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
});
