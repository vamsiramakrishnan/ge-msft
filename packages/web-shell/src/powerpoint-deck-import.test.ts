import { describe, expect, it } from 'vitest';
import { ActuationRequestSchema } from '@ge/contracts';
import { buildPowerPointDeckImportRequest } from './powerpoint-deck-import.js';

describe('buildPowerPointDeckImportRequest', () => {
  it('compiles a DeckSpec into one typed PowerPoint insert-slide deck actuation', async () => {
    const plan = await buildPowerPointDeckImportRequest({
      changeId: 'deck-test',
      formatting: 'UseDestinationTheme',
      targetSlideId: '256#1',
      deckSpec: {
        title: 'Generated deck',
        slides: [
          {
            title: 'Executive summary',
            elements: [
              {
                kind: 'text',
                text: 'The generated deck is inserted as one PPTX artifact.',
                x: 0.8,
                y: 1.4,
                w: 7,
                h: 0.7,
              },
            ],
          },
        ],
      },
    });

    expect(plan.artifact.slideCount).toBe(1);
    expect(plan.artifact.base64.startsWith('UEs')).toBe(true);
    expect(plan.request).toMatchObject({
      changeId: 'deck-test',
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: {
        deck: {
          format: 'pptx',
          slideCount: 1,
          formatting: 'UseDestinationTheme',
          targetSlideId: '256#1',
          specFingerprint: plan.artifact.specFingerprint,
        },
      },
    });
    expect(() => ActuationRequestSchema.parse(plan.request)).not.toThrow();
  });
});
