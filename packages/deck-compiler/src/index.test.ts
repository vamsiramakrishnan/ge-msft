import { describe, expect, it } from 'vitest';
import { compileDeckSpecToBase64, fingerprintDeckSpec } from './index.js';

describe('compileDeckSpecToBase64', () => {
  it('compiles a bounded deck spec into one base64 pptx artifact', async () => {
    const artifact = await compileDeckSpecToBase64({
      title: 'QBR',
      author: 'Test',
      slides: [
        {
          title: 'North region',
          subtitle: 'Generated from a validated deck spec',
          notes: 'Use this slide to explain the revenue bridge.',
          elements: [
            {
              kind: 'text',
              text: 'Revenue grew 12% while support backlog fell.',
              x: 0.8,
              y: 1.5,
              w: 6.2,
              h: 0.7,
            },
            {
              kind: 'bullets',
              items: ['ARR: $4.2M', 'NPS: 52', 'Open risks: 2'],
              x: 0.9,
              y: 2.4,
              w: 5.8,
              h: 1.2,
            },
          ],
        },
        {
          title: 'Pipeline',
          elements: [
            {
              kind: 'table',
              rows: [
                ['Stage', 'Count'],
                ['Qualified', '18'],
                ['Commit', '6'],
              ],
              x: 0.8,
              y: 1.4,
              w: 5.5,
              h: 1.4,
            },
          ],
        },
      ],
    });
    expect(artifact.format).toBe('pptx');
    expect(artifact.slideCount).toBe(2);
    expect(artifact.base64.startsWith('UEs')).toBe(true);
    expect(artifact.specFingerprint).toMatch(/^[0-9a-f]{8}$/);
  });

  it('reports adjusted frames instead of emitting invalid slide geometry', async () => {
    const artifact = await compileDeckSpecToBase64({
      slides: [
        {
          elements: [{ kind: 'text', text: 'Too far right', x: 15, y: 1, w: 10, h: 1 }],
        },
      ],
    });
    expect(artifact.warnings).toContain(
      'Adjusted an element frame to fit within 13.333x7.5 slides.',
    );
  });
});

describe('fingerprintDeckSpec', () => {
  it('is deterministic across object key order', () => {
    const a = fingerprintDeckSpec({
      title: 'A',
      layout: 'wide',
      slides: [{ title: 'One', elements: [] }],
    });
    const b = fingerprintDeckSpec({
      layout: 'wide',
      slides: [{ title: 'One', elements: [] }],
      title: 'A',
    });
    expect(a).toBe(b);
  });
});
