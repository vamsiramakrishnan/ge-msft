import { describe, expect, it } from 'vitest';
import { bindBoardPackTemplate, compileBoardPack, type BoardPackTemplate } from './board-pack.js';
import { compileDeckSpecToBase64 } from './index.js';

const id = 'a_' + 'a'.repeat(24);
const artifact = {
  id,
  hash: 'sha256:' + 'b'.repeat(64),
  title: 'Revenue',
  createdAt: '2026-10-25T00:00:00Z',
  columns: [
    { name: 'c0', label: 'Region', type: 'string' },
    { name: 'c1', label: 'Revenue', type: 'string' },
  ],
  rows: [
    ['East', '1234.50'],
    ['West', '56.78'],
  ],
  sources: [],
  lineage: { parents: [], operation: 'query' },
  truncated: false,
};
const template: BoardPackTemplate = {
  version: 1,
  spec: {
    masters: [
      {
        name: 'Brand',
        elements: [{ kind: 'text', text: 'Confidential', x: 0.5, y: 7, w: 3, h: 0.25 }],
      },
    ],
    slides: [
      {
        id: 'summary',
        master: 'Brand',
        elements: [
          { kind: 'text', text: 'Revenue {{revenue}}', x: 1, y: 1, w: 9, h: 0.5 },
          {
            kind: 'chart',
            chartType: 'column',
            labels: ['Pending'],
            series: [{ name: 'Revenue', values: [0] }],
            x: 1,
            y: 2,
            w: 9,
            h: 4,
          },
          { kind: 'text', text: 'User commentary', x: 10.5, y: 2, w: 2, h: 2 },
        ],
      },
    ],
  },
  bindings: [
    {
      kind: 'text',
      slideId: 'summary',
      element: 0,
      values: { revenue: { artifactId: id, row: 0, column: 1 } },
    },
    {
      kind: 'chart',
      slideId: 'summary',
      element: 1,
      artifactId: id,
      rows: [0, 1],
      labelColumn: 0,
      series: [{ name: 'Revenue', column: 1 }],
    },
  ],
};

describe('template-aware board pack compiler', () => {
  it('emits reusable native master and editable chart with exact source lineage', async () => {
    const result = await compileBoardPack(template, [artifact], {
      compression: false,
      strictLayout: true,
    });
    expect(result.spec.slides[0]?.elements[0]).toMatchObject({ text: 'Revenue 1234.50' });
    expect(result.spec.slides[0]?.elements[1]).toMatchObject({
      labels: ['East', 'West'],
      series: [{ name: 'Revenue', values: [1234.5, 56.78] }],
    });
    expect(result.lineage[0]).toMatchObject({
      artifactId: id,
      artifactHash: artifact.hash,
      cells: expect.arrayContaining([{ row: 0, column: 1, exactValue: '1234.50' }]),
    });
    const zip = Buffer.from(result.artifact.base64, 'base64').toString('utf8');
    expect(zip).toContain('ppt/slideMasters/slideMaster');
    expect(zip).toContain('Confidential');
    expect(zip).toContain('c:barChart');
    expect(zip).toContain('ppt/embeddings/');
  });
  it('refreshes bound metrics/charts while preserving unbound manual edits', () => {
    const previous = bindBoardPackTemplate(template, [artifact]).spec;
    const current = structuredClone(previous);
    current.slides[0]!.elements[2] = {
      kind: 'text',
      text: 'Manual commentary updated',
      x: 10.5,
      y: 2,
      w: 2,
      h: 2,
    };
    const next = bindBoardPackTemplate(
      template,
      [
        {
          ...artifact,
          rows: [
            ['East', '2000.00'],
            ['West', '60.25'],
          ],
        },
      ],
      current,
      previous,
    );
    expect(next.spec.slides[0]?.elements[0]).toMatchObject({ text: 'Revenue 2000.00' });
    expect(next.spec.slides[0]?.elements[2]).toMatchObject({ text: 'Manual commentary updated' });
  });
  it('refuses a manually changed managed chart and leaves the input untouched', () => {
    const previous = bindBoardPackTemplate(template, [artifact]).spec;
    const current = structuredClone(previous);
    const element = current.slides[0]!.elements[1]!;
    if (element.kind !== 'chart') throw new Error('fixture');
    element.series[0]!.values[0] = 999;
    expect(() => bindBoardPackTemplate(template, [artifact], current, previous)).toThrow(
      'manual edits',
    );
    expect(element.series[0]!.values[0]).toBe(999);
  });
  it('refuses incomplete artifacts, missing cells, ambiguous ids and unsafe chart precision', () => {
    expect(() => bindBoardPackTemplate(template, [{ ...artifact, truncated: true }])).toThrow(
      'complete',
    );
    expect(() => bindBoardPackTemplate(template, [{ ...artifact, rows: [] }])).toThrow('missing');
    expect(() => bindBoardPackTemplate(template, [artifact, artifact])).toThrow('uniquely');
    expect(() =>
      bindBoardPackTemplate(template, [
        {
          ...artifact,
          rows: [
            ['East', '123456789012345678.90'],
            ['West', '1'],
          ],
        },
      ]),
    ).toThrow('represented safely');
  });
  it('preserves long exact decimal text where no chart conversion is requested', () => {
    const onlyText = { ...template, bindings: [template.bindings[0]!] };
    const bound = bindBoardPackTemplate(onlyText, [
      {
        ...artifact,
        rows: [
          ['East', '123456789012345678.90'],
          ['West', '1'],
        ],
      },
    ]);
    expect(bound.spec.slides[0]?.elements[0]).toMatchObject({
      text: 'Revenue 123456789012345678.90',
    });
  });
  it('rejects unknown or duplicate masters and master content collisions', async () => {
    await expect(
      compileDeckSpecToBase64({ slides: [{ master: 'Unknown', elements: [] }] }),
    ).rejects.toThrow('unknown');
    await expect(
      compileDeckSpecToBase64({
        masters: [{ name: 'A' }, { name: 'A' }],
        slides: [{ elements: [] }],
      }),
    ).rejects.toThrow('Duplicate');
    await expect(
      compileDeckSpecToBase64(
        {
          masters: [
            { name: 'A', elements: [{ kind: 'text', text: 'Brand', x: 1, y: 1, w: 2, h: 1 }] },
          ],
          slides: [
            { master: 'A', elements: [{ kind: 'text', text: 'Content', x: 1, y: 1, w: 2, h: 1 }] },
          ],
        },
        { strictLayout: true },
      ),
    ).rejects.toThrow('overlaps');
  });
});
