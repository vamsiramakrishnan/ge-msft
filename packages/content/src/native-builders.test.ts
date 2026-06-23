import { describe, it, expect } from 'vitest';
import { native } from './index.js';
import { chunkBlocks } from './chunk.js';
import { processNative, toContextNative } from './process.js';
import type { NativeContent } from './model.js';

/**
 * Direct coverage of the native Block builders the bridges use to turn the Office object
 * model into Block[] without a Markdown round-trip. These assert the exact Markdown text
 * each builder emits and that the host locator is threaded through (or omitted) verbatim.
 */
describe('native block builders', () => {
  describe('heading', () => {
    it('prefixes the text with the right number of # for the level', () => {
      expect(native.heading('Intro', 1).text).toBe('# Intro');
      expect(native.heading('Sub', 3).text).toBe('### Sub');
      expect(native.heading('Deep', 6).text).toBe('###### Deep');
    });

    it('clamps the level into the 1..6 ATX range', () => {
      // level 0 and below clamp up to 1
      expect(native.heading('Zero', 0).text).toBe('# Zero');
      expect(native.heading('Neg', -5).text).toBe('# Neg');
      // level above 6 clamps down to 6
      expect(native.heading('Huge', 12).text).toBe('###### Huge');
    });

    it('keeps the raw numeric level field as supplied (not the clamped one)', () => {
      const h = native.heading('Intro', 9);
      expect(h.kind).toBe('heading');
      expect(h.level).toBe(9);
    });

    it('attaches the locator only when provided', () => {
      expect(native.heading('A', 1, 'cc:7').locator).toBe('cc:7');
      expect('locator' in native.heading('A', 1)).toBe(false);
    });
  });

  describe('paragraph', () => {
    it('passes text through unchanged and carries the locator', () => {
      const p = native.paragraph('Hello world.', 'cc:1');
      expect(p).toEqual({ kind: 'paragraph', text: 'Hello world.', locator: 'cc:1' });
    });

    it('omits the locator key when none is given', () => {
      const p = native.paragraph('No anchor');
      expect(p).toEqual({ kind: 'paragraph', text: 'No anchor' });
      expect('locator' in p).toBe(false);
    });
  });

  describe('listBlock', () => {
    it('renders items as a dash bullet list joined by newlines', () => {
      const b = native.listBlock(['one', 'two', 'three'], 'cc:5');
      expect(b.kind).toBe('list');
      expect(b.text).toBe('- one\n- two\n- three');
      expect(b.locator).toBe('cc:5');
    });

    it('handles an empty item array as an empty body', () => {
      const b = native.listBlock([]);
      expect(b.text).toBe('');
      expect('locator' in b).toBe(false);
    });
  });

  describe('quote', () => {
    it('prefixes every line of a multi-line quote with "> "', () => {
      const b = native.quote('first\nsecond', 'cc:9');
      expect(b.kind).toBe('quote');
      expect(b.text).toBe('> first\n> second');
      expect(b.locator).toBe('cc:9');
    });

    it('prefixes a single line and omits an absent locator', () => {
      const b = native.quote('alone');
      expect(b.text).toBe('> alone');
      expect('locator' in b).toBe(false);
    });
  });

  describe('code', () => {
    it('wraps the text in a fenced code block', () => {
      const b = native.code('const x = 1;', 'cc:3');
      expect(b.kind).toBe('code');
      expect(b.text).toBe('```\nconst x = 1;\n```');
      expect(b.locator).toBe('cc:3');
    });

    it('omits the locator key when not provided', () => {
      const b = native.code('noop()');
      expect('locator' in b).toBe(false);
    });
  });

  describe('table', () => {
    it('renders structured data to GFM and preserves the structured data + locator', () => {
      const data = { columns: ['A', 'B'], rows: [['1', '2'] as (string | number)[]] };
      const b = native.table(data, 'range:Sheet1!A1:B2');
      expect(b.kind).toBe('table');
      expect(b.text).toContain('| A | B |');
      expect(b.text).toContain('| 1 | 2 |');
      expect(b.data).toBe(data);
      expect(b.locator).toBe('range:Sheet1!A1:B2');
    });

    it('omits the locator key when not provided but still keeps data', () => {
      const b = native.table({ columns: ['X'], rows: [] });
      expect('locator' in b).toBe(false);
      expect(b.data).toEqual({ columns: ['X'], rows: [] });
    });
  });

  describe('slide', () => {
    it('uses an explicit slideId locator and emits title heading + each non-empty body para', () => {
      const blocks = native.slide(2, 'Risk', ['Line one.', 'Line two.'], 'sid-7');
      expect(blocks).toHaveLength(3);
      expect(blocks[0]).toMatchObject({ kind: 'heading', level: 2, locator: 'slide:sid-7' });
      expect(blocks[0]!.text).toBe('## Risk');
      expect(blocks[1]).toMatchObject({
        kind: 'paragraph',
        text: 'Line one.',
        locator: 'slide:sid-7',
      });
      expect(blocks[2]!.text).toBe('Line two.');
    });

    it('falls back to a slide:index locator when no slideId is given', () => {
      const blocks = native.slide(4, 'T', []);
      expect(blocks[0]!.locator).toBe('slide:4');
    });

    it('synthesizes "Slide N+1" when the title is empty', () => {
      const blocks = native.slide(0, '', ['body']);
      expect(blocks[0]!.text).toBe('## Slide 1');
    });

    it('drops blank/whitespace-only body paragraphs', () => {
      const blocks = native.slide(0, 'Title', ['  ', '', 'real', '\t\n']);
      const paras = blocks.filter((b) => b.kind === 'paragraph');
      expect(paras).toHaveLength(1);
      expect(paras[0]!.text).toBe('real');
    });

    it('produces only the title heading for an all-empty body', () => {
      const blocks = native.slide(1, 'Solo', []);
      expect(blocks).toHaveLength(1);
      expect(blocks[0]!.kind).toBe('heading');
    });
  });
});

describe('native builders feed the chunk pipeline', () => {
  it('a list + quote + code block flow through processNative into one anchored chunk', () => {
    const doc: NativeContent = {
      sourceId: 'word:body',
      surface: 'word',
      blocks: [
        native.heading('Notes', 2, 'cc:1'),
        native.listBlock(['a', 'b'], 'cc:2'),
        native.quote('cited', 'cc:3'),
        native.code('x()', 'cc:4'),
      ],
    };
    const { chunks } = processNative(doc, { maxTokens: 400 });
    expect(chunks).toHaveLength(1);
    const c = chunks[0]!;
    // The native locator of the first block in the chunk wins as the anchor.
    expect(c.meta.anchor.locator).toBe('cc:1');
    expect(c.meta.kinds).toEqual(expect.arrayContaining(['heading', 'list', 'quote', 'code']));
    expect(c.text).toContain('- a');
    expect(c.text).toContain('> cited');
    expect(c.text).toContain('```');
  });

  it('toContextNative without a surface defaults the ref surface to word', () => {
    const doc: NativeContent = {
      sourceId: 'src:1',
      blocks: [native.paragraph('plain text body', 'cc:0')],
    };
    const ctx = toContextNative(doc);
    expect(ctx).toHaveLength(1);
    expect(ctx[0]!.ref.surface).toBe('word');
  });

  it('toContextNative honours preferReference + indexedDocumentName (reference over inline)', () => {
    const doc: NativeContent = {
      sourceId: 'src:9',
      surface: 'powerpoint',
      title: 'Deck',
      indexedDocumentName: 'projects/x/dataStores/d/documents/9',
      blocks: [native.paragraph('ignored when referenced')],
    };
    const ctx = toContextNative(doc, { preferReference: true });
    expect(ctx).toHaveLength(1);
    expect(ctx[0]!.value).toMatchObject({
      as: 'indexed-document',
      documentName: 'projects/x/dataStores/d/documents/9',
      title: 'Deck',
    });
    expect(ctx[0]!.ref.surface).toBe('powerpoint');
  });

  it('does not reference when preferReference is set but the source is not indexed', () => {
    const doc: NativeContent = {
      sourceId: 'src:10',
      surface: 'excel',
      blocks: [native.paragraph('inline body wins')],
    };
    const ctx = toContextNative(doc, { preferReference: true });
    expect(ctx[0]!.value).toMatchObject({ as: 'text', mimeType: 'text/markdown' });
  });

  it('chunkBlocks on an empty native block list yields no chunks', () => {
    expect(chunkBlocks([], { sourceId: 'empty:1' })).toEqual([]);
  });
});
