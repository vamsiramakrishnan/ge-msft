import { describe, it, expect, afterEach } from 'vitest';
import { PowerPointBridge } from '@ge/bridge-powerpoint';
import { ActuationRequestSchema, type ActuationRequest } from '@ge/contracts';
import {
  installFakePowerPoint,
  powerPointSeed,
  type PowerPointSimulator,
} from './fake-powerpoint.js';

/**
 * The REAL `PowerPointBridge` shape writes against the PowerPoint simulator: each lands `ok: true`
 * and the deck snapshot shows the change.
 */

let sim: PowerPointSimulator | undefined;

afterEach(() => {
  sim?.restore();
  sim = undefined;
});

function request(kind: string, params: Record<string, unknown>): ActuationRequest {
  return ActuationRequestSchema.parse({
    changeId: `c-${kind}`,
    kind,
    surface: 'powerpoint',
    params,
  });
}

function install(): PowerPointSimulator {
  sim = installFakePowerPoint(
    powerPointSeed({
      slides: [
        {
          id: 'slide-1',
          shapes: [
            { text: 'Q3 Business Review', type: 'Placeholder', placeholderType: 'Title' },
            { text: 'Acme Corp', type: 'Placeholder', placeholderType: 'Body' },
          ],
        },
        { id: 'slide-2', shapes: [{ text: 'Revenue' }, { text: 'Up 12% QoQ' }] },
      ],
    }),
  );
  return sim;
}

describe('fake PowerPoint: shape writes through the real bridge', () => {
  it('set-shape-text rewrites the title shape', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('set-shape-text', {
        target: { slideId: 'slide-1', shapeId: 'title' },
        text: 'FY26 Plan',
      }),
    );
    expect(result).toMatchObject({ ok: true, location: 'shape:slide-1:slide-1-shape-0' });
    expect(result.inverse).toMatchObject({ op: 'restore-text', priorText: 'Q3 Business Review' });
    expect(deck.snapshot().slides[0]?.shapeTexts[0]).toBe('FY26 Plan');
  });

  it('set-shape-text with numbered lines lands and hides the layout bullet', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('set-shape-text', {
        target: { slideId: 'slide-1', shapeId: 'slide-1-shape-1' },
        text: '1. Ringkasan Eksekutif\n2. Analisa Kebutuhan\n3. Roadmap',
      }),
    );
    expect(result).toMatchObject({ ok: true });
    expect(deck.snapshot().slides[0]?.shapeTexts[1]).toBe(
      '1. Ringkasan Eksekutif\r2. Analisa Kebutuhan\r3. Roadmap',
    );
    expect(deck.seed.slides[0]?.shapes[1]?.bulletVisible).toBe(false);
  });

  it('set-shape-text on a missing shape id degrades with the valid choices', async () => {
    install();
    const result = await new PowerPointBridge().actuate(
      request('set-shape-text', { target: { slideId: '1', shapeId: 'nope' }, text: 'x' }),
    );
    expect(result.ok).toBe(false);
    expect(result.error?.message).toContain('slide-1-shape-0');
  });

  it('add-shape adds a positioned text box with a minted id', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('add-shape', {
        target: { slideId: '1' },
        shape: {
          shapeType: 'textBox',
          text: 'Draft for review',
          left: 72,
          top: 300,
          width: 400,
          height: 50,
        },
      }),
    );
    expect(result).toMatchObject({ ok: true, location: 'shape:slide-1:slide-1-shape-2' });
    expect(deck.snapshot().slides[0]?.shapes[2]).toMatchObject({
      id: 'slide-1-shape-2',
      type: 'TextBox',
      text: 'Draft for review',
      left: 72,
      top: 300,
      width: 400,
      height: 50,
    });
  });

  it('add-shape adds a filled geometric shape', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('add-shape', {
        target: { slideId: 'slide-2' },
        shape: {
          shapeType: 'geometric',
          geometryType: 'Ellipse',
          fill: '#FF0000',
          left: 10,
          top: 10,
          width: 50,
          height: 50,
        },
      }),
    );
    expect(result.ok).toBe(true);
    expect(deck.snapshot().slides[1]?.shapes.at(-1)).toMatchObject({
      type: 'GeometricShape',
      geometricShapeType: 'Ellipse',
      fill: '#FF0000',
    });
  });

  it('format-shape sets bold, colour, size and fill on the title', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('format-shape', {
        target: { slideId: '1', shapeId: 'title' },
        shapeFormat: { fill: '#112233', font: { bold: true, color: '#C00000', size: 40 } },
      }),
    );
    expect(result).toMatchObject({ ok: true, location: 'shape:slide-1:slide-1-shape-0' });
    expect(result.inverse).toMatchObject({
      op: 'restore-shape-format',
      shapeId: 'slide-1-shape-0',
    });
    expect(deck.snapshot().slides[0]?.shapes[0]).toMatchObject({
      fill: '#112233',
      font: { bold: true, color: '#C00000', size: 40 },
    });
  });

  it('add-table-slide creates a titled slide holding the table', async () => {
    const deck = install();
    const result = await new PowerPointBridge().actuate(
      request('add-table-slide', {
        target: { slideId: 'new' },
        slide: { title: 'Regional revenue', bullets: [] },
        tableGrid: {
          rows: [
            ['Region', 'Revenue'],
            ['East', '12'],
            ['West', '9'],
          ],
        },
      }),
    );
    expect(result.ok).toBe(true);
    const slides = deck.snapshot().slides;
    expect(slides).toHaveLength(3);
    const added = slides[2]!;
    expect(result.location).toBe(`shape:${added.id}:${added.id}-shape-2`);
    expect(result.inverse).toMatchObject({
      op: 'delete-object',
      objectType: 'slide',
      name: added.id,
    });
    expect(added.shapeTexts[0]).toBe('Regional revenue');
    expect(added.shapes[2]).toMatchObject({
      type: 'Table',
      table: [
        ['Region', 'Revenue'],
        ['East', '12'],
        ['West', '9'],
      ],
    });
  });

  it('a read after the writes sees the new table text', async () => {
    install();
    const bridge = new PowerPointBridge();
    await bridge.actuate(
      request('add-table-slide', {
        target: { slideId: '2' },
        tableGrid: { rows: [['Churn', '3%']], left: 10, top: 10, width: 200, height: 40 },
      }),
    );
    const found = await bridge.searchDocument('Churn');
    expect(found.length).toBeGreaterThan(0);
  });
});
