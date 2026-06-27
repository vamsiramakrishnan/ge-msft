import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import { native, processNative, toContextNative } from './index.js';
import type { NativeContent } from './model.js';

/**
 * Simulates what the bridges produce from the Office object model — no Markdown
 * round-trip, native host locators (content-control id, range address, slide index)
 * preserved straight through to the write-back anchor.
 */
describe('native path — Word object model', () => {
  const doc: NativeContent = {
    sourceId: 'word:body',
    title: 'MSA',
    surface: 'word',
    blocks: [
      native.heading('5. Service Levels', 1, 'cc:10'),
      native.heading('5.1 Availability', 2, 'cc:11'),
      native.paragraph('The services are available 99.5% of the time.', 'cc:12'),
    ],
  };

  it('keeps the content-control id as the anchor locator', () => {
    const { chunks } = processNative(doc, { maxTokens: 40 });
    const avail = chunks.find((c) => c.text.includes('99.5%'))!;
    expect(avail.meta.sectionPath).toEqual(['5. Service Levels', '5.1 Availability']);
    expect(avail.meta.anchor.locator).toBe('cc:11'); // first block in the chunk (the heading)
    expect(avail.meta.anchor.contextHint).toContain('Availability');
    expect(avail.meta.charStart).toBeUndefined(); // native path: no char offsets
  });

  it('emits valid attach-ready context', () => {
    for (const c of toContextNative(doc, { maxTokens: 40 })) {
      expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
      expect(c.value).toMatchObject({ as: 'text', mimeType: 'text/markdown' });
    }
  });
});

describe('native path — Excel range as structured data', () => {
  const sheet: NativeContent = {
    sourceId: 'xl:Sheet1!A1:B3',
    title: 'Vendors',
    surface: 'excel',
    blocks: [
      native.table(
        {
          columns: ['Vendor', 'Risk'],
          rows: [
            ['Acme', 'High'],
            ['Globex', 'Low'],
          ],
        },
        'range:Sheet1!A1:B3',
      ),
    ],
  };

  it('renders the range as one GFM table chunk anchored to its address', () => {
    const { chunks } = processNative(sheet, { maxTokens: 5 }); // tiny budget: still not split
    expect(chunks).toHaveLength(1);
    expect(chunks[0]!.text).toContain('| Vendor | Risk |');
    expect(chunks[0]!.text).toContain('| Acme | High |');
    expect(chunks[0]!.meta.anchor.locator).toBe('range:Sheet1!A1:B3');
    expect(chunks[0]!.meta.kinds).toContain('table');
  });
});

describe('native path — PowerPoint slide', () => {
  it('turns a slide into title + body anchored to the slide', () => {
    const deck: NativeContent = {
      sourceId: 'ppt:deck',
      surface: 'powerpoint',
      blocks: native.slide(3, 'Risk Summary', ['Vendor exposure is elevated.'], 'slide-id-9'),
    };
    const { chunks } = processNative(deck);
    expect(chunks[0]!.text).toContain('## Risk Summary');
    expect(chunks[0]!.meta.anchor.locator).toBe('slide:slide-id-9');
  });
});
