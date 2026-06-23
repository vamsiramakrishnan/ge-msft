import { describe, it, expect } from 'vitest';
import { parseMarkdownBlocks, tableToMarkdown } from './markdown.js';

/**
 * Exercises the line-based Markdown block parser across every block kind, with emphasis on
 * the fenced-code, blockquote, and list branches (offset correctness, run consumption, and
 * the degenerate inputs the existing suite did not cover).
 */
describe('parseMarkdownBlocks — fenced code', () => {
  it('captures a fenced block as one code block with offsets spanning the source', () => {
    const md = '```\nline a\nline b\n```';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('code');
    expect(md.slice(blocks[0]!.start, blocks[0]!.end)).toContain('line a');
    expect(md.slice(blocks[0]!.start, blocks[0]!.end)).toContain('line b');
  });

  it('treats ~~~ as a fence too and does not parse structure inside the fence', () => {
    const md = '~~~\n# not a heading\n- not a list\n~~~';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('code');
    // The inner "# not a heading" must remain inside the single code block, not become a heading.
    expect(blocks.filter((b) => b.kind === 'heading')).toHaveLength(0);
    expect(blocks[0]!.text).toContain('# not a heading');
  });

  it('handles an unterminated fence by consuming to end of input', () => {
    const md = '```\nunclosed body\nstill code';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('code');
    expect(blocks[0]!.text).toContain('still code');
  });
});

describe('parseMarkdownBlocks — blockquote', () => {
  it('groups consecutive > lines into a single quote block', () => {
    const md = '> first\n> second\n> third';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('quote');
    expect(blocks[0]!.text).toContain('first');
    expect(blocks[0]!.text).toContain('third');
  });

  it('ends the quote at the first non-quote line', () => {
    const md = '> quoted\nplain paragraph after';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks.map((b) => b.kind)).toEqual(['quote', 'paragraph']);
    expect(blocks[1]!.text).toBe('plain paragraph after');
  });
});

describe('parseMarkdownBlocks — lists', () => {
  it('recognizes dash, star, plus, and ordered markers as one list run', () => {
    const md = '- a\n* b\n+ c\n1. d\n2) e';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('list');
    expect(blocks[0]!.text).toContain('a');
    expect(blocks[0]!.text).toContain('e');
  });

  it('absorbs indented continuation lines into the same list block', () => {
    const md = '- item one\n  continued text\n- item two';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks).toHaveLength(1);
    expect(blocks[0]!.kind).toBe('list');
    expect(blocks[0]!.text).toContain('continued text');
  });

  it('ends the list at a blank line', () => {
    const md = '- only item\n\na paragraph';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks.map((b) => b.kind)).toEqual(['list', 'paragraph']);
  });
});

describe('parseMarkdownBlocks — headings and degenerate input', () => {
  it('records the heading level from the # count', () => {
    const blocks = parseMarkdownBlocks('### Third level');
    expect(blocks[0]!.kind).toBe('heading');
    expect(blocks[0]!.level).toBe(3);
  });

  it('skips blank and whitespace-only lines, producing no empty blocks', () => {
    const blocks = parseMarkdownBlocks('\n\n   \n\n');
    expect(blocks).toHaveLength(0);
  });

  it('returns an empty array for an empty string', () => {
    expect(parseMarkdownBlocks('')).toEqual([]);
  });

  it('terminates a paragraph at a following structural (heading) line', () => {
    const md = 'para line one\npara line two\n# Heading';
    const blocks = parseMarkdownBlocks(md);
    expect(blocks.map((b) => b.kind)).toEqual(['paragraph', 'heading']);
    expect(blocks[0]!.text).toBe('para line one\npara line two');
  });

  it('parses a full mixed document into the expected kind sequence', () => {
    const md = [
      '# Title',
      '',
      'Intro paragraph.',
      '',
      '- bullet',
      '',
      '> quoted',
      '',
      '```',
      'code()',
      '```',
      '',
      '| H |',
      '| --- |',
      '| v |',
    ].join('\n');
    const kinds = parseMarkdownBlocks(md).map((b) => b.kind);
    expect(kinds).toEqual(['heading', 'paragraph', 'list', 'quote', 'code', 'table']);
  });
});

describe('tableToMarkdown', () => {
  it('builds the header, separator, and body rows; coerces numbers to strings', () => {
    const md = tableToMarkdown(
      ['Q', 'N'],
      [
        ['x', 1 as unknown as number],
        [2, 3],
      ],
    );
    const lines = md.split('\n');
    expect(lines[0]).toBe('| Q | N |');
    expect(lines[1]).toBe('| --- | --- |');
    expect(lines[2]).toBe('| x | 1 |');
    expect(lines[3]).toBe('| 2 | 3 |');
  });

  it('emits an empty body line when there are no rows', () => {
    const md = tableToMarkdown(['Only'], []);
    expect(md).toBe('| Only |\n| --- |\n');
  });
});
