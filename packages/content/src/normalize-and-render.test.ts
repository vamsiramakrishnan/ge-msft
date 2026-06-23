import { describe, it, expect } from 'vitest';
import { toMarkdown, htmlToMarkdown } from './normalize.js';
import { buildDocStateSnapshot, renderDocState } from './doc-state-builder.js';
import type { RawContent } from './model.js';

const FIXED_NOW = (): Date => new Date('2026-06-22T12:00:00.000Z');

/**
 * toMarkdown dispatches on the declared format. Each branch (markdown passthrough, plain
 * text, html reduction) is exercised, plus the html reducer's security-relevant stripping.
 */
describe('toMarkdown format dispatch', () => {
  const raw = (text: string, format: RawContent['format']): RawContent => ({
    sourceId: 'src',
    text,
    format,
  });

  it('passes markdown through, trimmed', () => {
    expect(toMarkdown(raw('  # Heading\n\nbody  ', 'markdown'))).toBe('# Heading\n\nbody');
  });

  it('passes plain text through, trimmed', () => {
    expect(toMarkdown(raw('\n\nplain body line\n', 'plain'))).toBe('plain body line');
  });

  it('reduces html to markdown via htmlToMarkdown, trimmed', () => {
    const out = toMarkdown(raw('<h1>Title</h1><p>Body</p>', 'html'));
    expect(out).toContain('# Title');
    expect(out).toContain('Body');
    expect(out.startsWith('\n')).toBe(false); // trimmed
  });
});

describe('htmlToMarkdown — untrusted markup reduction', () => {
  it('drops script and style content entirely (no executable leakage)', () => {
    const out = htmlToMarkdown('<style>.a{color:red}</style><script>alert(1)</script><p>safe</p>');
    expect(out).not.toContain('alert(1)');
    expect(out).not.toContain('color:red');
    expect(out).toContain('safe');
  });

  it('converts headings, lists, emphasis, and breaks', () => {
    const out = htmlToMarkdown(
      '<h3>Sec</h3><ul><li>one</li><li>two</li></ul><p>A<br>B <strong>bold</strong> <em>it</em></p>',
    );
    expect(out).toContain('### Sec');
    expect(out).toContain('- one');
    expect(out).toContain('- two');
    expect(out).toContain('**bold**');
    expect(out).toContain('*it*');
  });

  it('decodes HTML entities', () => {
    const out = htmlToMarkdown('<p>a &amp; b &lt;tag&gt; &quot;q&quot; &#39;s&#39; &nbsp;x</p>');
    expect(out).toContain('a & b <tag> "q" \'s\'');
  });

  it('collapses 3+ blank lines down to a paragraph break', () => {
    const out = htmlToMarkdown('<div>one</div><div></div><div></div><div>two</div>');
    expect(out).not.toMatch(/\n{3,}/);
  });

  it('strips unknown/unhandled tags but keeps their text', () => {
    const out = htmlToMarkdown('<span class="x">kept text</span>');
    expect(out).toContain('kept text');
    expect(out).not.toContain('<span');
  });
});

/**
 * renderDocState has a namedRanges section only reached when the snapshot carries ranges;
 * the existing render tests never feed one in.
 */
describe('renderDocState — named ranges section', () => {
  it('renders each named range as a name = range line', () => {
    const snap = buildDocStateSnapshot({
      surface: 'excel',
      version: 1,
      blocks: [],
      namedRanges: [
        { name: 'Revenue', range: 'Sheet1!$A$1:$A$12' },
        { name: 'Costs', range: 'Sheet1!$B$1:$B$12' },
      ],
      now: FIXED_NOW,
    });
    const out = renderDocState(snap);
    expect(out).toContain('namedRanges:');
    expect(out).toContain('- "Revenue" = "Sheet1!$A$1:$A$12"');
    expect(out).toContain('- "Costs" = "Sheet1!$B$1:$B$12"');
  });

  it('omits the namedRanges section entirely when there are none', () => {
    const snap = buildDocStateSnapshot({
      surface: 'excel',
      version: 1,
      blocks: [],
      now: FIXED_NOW,
    });
    expect(renderDocState(snap)).not.toContain('namedRanges:');
  });
});
