import { describe, expect, it } from 'vitest';
import { compileDeckSpecToBase64, inspectDeckLayout } from './index.js';

const chart = {
  kind: 'chart',
  chartType: 'column',
  labels: ['East', 'West'],
  series: [{ name: 'Revenue', values: [12.34, 56.78] }],
  x: 1,
  y: 1.5,
  w: 9,
  h: 4,
};
describe('editable chart artifacts and layout gates', () => {
  it('emits native chart XML and an embedded workbook rather than a chart image', async () => {
    const result = await compileDeckSpecToBase64(
      {
        theme: { headFontFace: 'Arial', bodyFontFace: 'Arial', accentColors: ['0057B8'] },
        slides: [{ title: 'Revenue', elements: [chart] }],
      },
      { compression: false, strictLayout: true },
    );
    const zip = Buffer.from(result.base64, 'base64').toString('utf8');
    expect(zip).toContain('ppt/charts/chart1.xml');
    expect(zip).toContain('c:barChart');
    expect(zip).toContain('ppt/embeddings/');
    expect(zip).toContain('56.78');
    expect(zip).toContain('0057B8');
    expect(zip).not.toContain('ppt/media/image');
  });
  it('rejects mismatched series and enforces aggregate asset budgets', async () => {
    await expect(
      compileDeckSpecToBase64({
        slides: [{ elements: [{ ...chart, series: [{ name: 'Revenue', values: [1] }] }] }],
      }),
    ).rejects.toThrow('label count');
    await expect(
      compileDeckSpecToBase64({
        slides: [
          {
            elements: Array.from({ length: 3 }, () => ({
              kind: 'image',
              base64: 'a'.repeat(6_000_000),
              mimeType: 'image/png',
              x: 0,
              y: 0,
              w: 1,
              h: 1,
            })),
          },
        ],
      }),
    ).rejects.toThrow('aggregate');
  });
  it('detects collisions and permits explicitly layered artwork', async () => {
    const spec = {
      slides: [
        {
          title: 'Title',
          elements: [{ kind: 'text', text: 'Colliding', x: 1, y: 0.4, w: 4, h: 1 }],
        },
      ],
    };
    expect(inspectDeckLayout(spec)).toEqual([expect.stringContaining('overlaps')]);
    await expect(compileDeckSpecToBase64(spec, { strictLayout: true })).rejects.toThrow('overlaps');
    expect(
      inspectDeckLayout({
        slides: [{ elements: [{ ...spec.slides[0]!.elements[0], allowOverlap: true }] }],
      }),
    ).toEqual([]);
  });
});
