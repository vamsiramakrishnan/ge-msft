import { describe, it, expect } from 'vitest';
import { escapeHtml, citationTag, partToHtml, buildPageHtml } from './synthesis.js';

describe('onenote synthesis (pure)', () => {
  it('escapes HTML-significant characters so text is rendered as data', () => {
    expect(escapeHtml('<b>"a" & \'b\'</b>')).toBe(
      '&lt;b&gt;&quot;a&quot; &amp; &#39;b&#39;&lt;/b&gt;',
    );
  });

  it('renders a citation tag with the source title and optional locator', () => {
    expect(citationTag({ title: 'Risk Policy', locator: '§3.2' })).toBe(
      '<span data-ge-cite="1">[Risk Policy · §3.2]</span>',
    );
    expect(citationTag({ title: 'ISO cert' })).toBe('<span data-ge-cite="1">[ISO cert]</span>');
  });

  it('renders a part as an optional heading + a paragraph with a trailing citation', () => {
    const html = partToHtml({
      heading: 'Where it falls short',
      text: 'The SLA of 99.5% sits below the 99.9% standard.',
      source: { title: 'Risk Policy', locator: '§3.2' },
    });
    expect(html).toBe(
      '<h2>Where it falls short</h2><p>The SLA of 99.5% sits below the 99.9% standard. <span data-ge-cite="1">[Risk Policy · §3.2]</span></p>',
    );
  });

  it('omits the heading when blank and the citation when there is no source', () => {
    expect(partToHtml({ text: 'plain claim' })).toBe('<p>plain claim</p>');
  });

  it('escapes claim text injected with markup', () => {
    expect(partToHtml({ text: '<script>x</script>' })).toBe(
      '<p>&lt;script&gt;x&lt;/script&gt;</p>',
    );
  });

  it('joins multiple parts into one body', () => {
    const html = buildPageHtml([{ text: 'a' }, { text: 'b' }]);
    expect(html).toBe('<p>a</p><p>b</p>');
  });
});
