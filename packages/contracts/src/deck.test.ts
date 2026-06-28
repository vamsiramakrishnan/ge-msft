import { describe, expect, it } from 'vitest';
import { ActuationParamsSchema, DeckSpecSchema } from './index.js';

describe('DeckSpecSchema', () => {
  it('accepts a bounded editable deck spec', () => {
    const spec = DeckSpecSchema.parse({
      title: 'Quarterly update',
      slides: [
        {
          title: 'Revenue',
          elements: [
            { kind: 'text', text: 'North region led growth', x: 0.7, y: 1.2, w: 6, h: 0.5 },
            {
              kind: 'bullets',
              items: ['Revenue up 12%', 'Churn down 2 pts'],
              x: 0.9,
              y: 2,
              w: 7,
              h: 1.4,
            },
          ],
        },
      ],
    });
    expect(spec.layout).toBe('wide');
    expect(spec.slides[0]?.elements).toHaveLength(2);
  });

  it('rejects oversized generated decks', () => {
    const slides = Array.from({ length: 81 }, (_, i) => ({ title: `Slide ${i + 1}` }));
    expect(() => DeckSpecSchema.parse({ slides })).toThrow();
  });
});

describe('PowerPoint deck actuation params', () => {
  it('accepts an explicit base64 PPTX deck import artifact', () => {
    const params = ActuationParamsSchema.parse({
      deck: {
        base64: 'UEsDBBQ=',
        slideCount: 3,
        formatting: 'UseDestinationTheme',
        targetSlideId: '256#1',
      },
    });
    expect(params.deck?.format).toBe('pptx');
  });
});
