import { describe, expect, it } from 'vitest';
import { ActuationRequestSchema } from '@ge/contracts';
import {
  buildPowerPointDeckImportRequest,
  buildPowerPointLayoutRequest,
} from './powerpoint-deck-import.js';

describe('buildPowerPointDeckImportRequest', () => {
  it('uses a fresh operation ID for separately approved imports of the same spec', async () => {
    const deckSpec = { slides: [{ title: 'Same content', elements: [] }] };
    const first = await buildPowerPointDeckImportRequest({ deckSpec });
    const second = await buildPowerPointDeckImportRequest({ deckSpec });
    expect(first.artifact.specFingerprint).toBe(second.artifact.specFingerprint);
    expect(first.request.changeId).not.toBe(second.request.changeId);
  });
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

describe('bound board pack import path', () => {
  it('compiles bound exact metrics through the existing reviewed insert-slide request', async () => {
    const id = 'a_' + 'a'.repeat(24);
    const plan = await buildPowerPointDeckImportRequest({
      boardPack: {
        template: {
          version: 1,
          spec: {
            slides: [
              { id: 's', elements: [{ kind: 'text', text: '{{metric}}', x: 1, y: 1, w: 4, h: 1 }] },
            ],
          },
          bindings: [
            {
              kind: 'text',
              slideId: 's',
              element: 0,
              values: { metric: { artifactId: id, row: 0, column: 0 } },
            },
          ],
        },
        artifacts: [
          {
            id,
            title: 'Exact',
            hash: 'sha256:' + 'b'.repeat(64),
            createdAt: '2026-10-25T00:00:00Z',
            columns: [{ name: 'c0', label: 'Amount', type: 'string' }],
            rows: [['123456789012345678.90']],
            sources: [],
            lineage: { parents: [], operation: 'query' },
            truncated: false,
          },
        ],
      },
    });
    expect(plan.request.kind).toBe('insert-slide');
    expect(plan.boundSpec?.slides[0]?.elements[0]).toMatchObject({ text: '123456789012345678.90' });
    expect(plan.lineage?.[0]?.artifactId).toBe(id);
  });
  it('rejects ambiguous import source selection', async () => {
    await expect(buildPowerPointDeckImportRequest({})).rejects.toThrow('exactly one');
  });
});

describe('captured native layout request', () => {
  it('pins saved document and exact master/layout hash through the ordinary effect contract', () => {
    const capture = {
      version: 1,
      documentId: 'sha256:' + 'a'.repeat(64),
      complete: true,
      omissions: [],
      masters: [
        {
          id: 'm',
          name: 'Brand',
          layouts: [
            { id: 'l', name: 'Title', revision: 'sha256:' + 'b'.repeat(64), placeholders: [] },
          ],
        },
      ],
    };
    expect(
      buildPowerPointLayoutRequest(capture, { masterId: 'm', layoutId: 'l', slideId: 's' }),
    ).toMatchObject({
      kind: 'apply-slide-layout',
      params: {
        target: { slideId: 's', expectedDocumentId: capture.documentId },
        layout: {
          masterId: 'm',
          layoutId: 'l',
          expectedLayoutHash: capture.masters[0]!.layouts[0]!.revision,
        },
      },
    });
    expect(() =>
      buildPowerPointLayoutRequest(
        { ...capture, complete: false },
        { masterId: 'm', layoutId: 'l', slideId: 's' },
      ),
    ).toThrow('complete');
    expect(() =>
      buildPowerPointLayoutRequest(capture, { masterId: 'm', layoutId: 'missing', slideId: 's' }),
    ).toThrow('Choose one');
  });
});
