// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { MessageThread, type MessageThreadProps } from './MessageThread.js';
import type { ChatMessage } from '../../controller.js';

/**
 * Behavioral tests for the conversation thread. The load-bearing logic here is the untrusted-URI
 * gate (`safeHttpUri`): citation URIs are grounded, untrusted source material, so a `javascript:`,
 * `data:`, or malformed URI must NOT become an executable href — it is rendered inert. Also covers
 * the empty-state invitation, per-message error/cancelled rendering, and the citation popover.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(messages: ChatMessage[], props: Partial<MessageThreadProps> = {}): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(createElement(MessageThread, { messages, ...props }));
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

/** Open the (single) citation popover so the link / inert text is in the DOM. */
function openFirstCitation(): void {
  const cite = container.querySelector<HTMLButtonElement>('.cite-btn');
  act(() => cite?.click());
}

describe('MessageThread', () => {
  it('shows the grounded empty-state invitation when there are no messages', () => {
    render([]);
    expect(container.textContent).toContain('Ask about this document or selection');
    // The invitation is an empty-state plate, not a persisted conversation message.
    expect(container.querySelectorAll('.thread-empty').length).toBe(1);
    expect(container.querySelectorAll('.m').length).toBe(0);
  });

  it('shows surface-specific empty-state copy', () => {
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
    act(() => {
      root.render(createElement(MessageThread, { messages: [], surface: 'excel' }));
    });
    expect(container.textContent).toContain('Ask about this workbook, sheet, or range');
  });

  it('renders a citation with an http(s) link as a real, new-tab, noopener anchor', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        sources: [{ title: 'Policy', uri: 'https://example.com/policy', locator: '§3' }],
      },
    ]);
    openFirstCitation();
    const link = container.querySelector<HTMLAnchorElement>('a.cite-detail-link');
    expect(link).not.toBeNull();
    expect(link?.getAttribute('href')).toBe('https://example.com/policy');
    expect(link?.getAttribute('target')).toBe('_blank');
    expect(link?.getAttribute('rel')).toBe('noreferrer noopener');
  });

  it('renders read and grid command payloads as compact cards, not raw TSV walls', () => {
    render([
      {
        id: 'u-1',
        role: 'user',
        text:
          "read 'Daily schedule'!A1:J54\n" +
          'grid \'Daily schedule\'!C7:I9 = "Music Lesson\\tIndia Sync\\tIndia Sync\\tIndia Sync\\t\\t\\tIndia Sync\\nCommute\\tIndia Sync\\tIndia Sync\\tIndia Sync\\t\\t\\tIndia Sync\\nDeep Work\\tDeep Work\\tDeep Work\\tDeep Work\\t\\t\\tDeep Work"',
      },
    ]);

    const read = container.querySelector('.cmd-card-read');
    const grid = container.querySelector('.cmd-card-grid');
    expect(read).not.toBeNull();
    expect(read?.textContent).toContain("'Daily schedule'!A1:J54");
    expect(grid).not.toBeNull();
    expect(grid?.textContent).toContain("'Daily schedule'!C7:I9");
    expect(grid?.textContent).toContain('3 x 7 cells');
    expect(grid?.querySelectorAll('.cmd-card-cell').length).toBe(12);
    expect(grid?.textContent).toContain('3 more columns hidden');
    // The escaped command body is summarized into cells rather than rendered as one raw line.
    expect(container.textContent).not.toContain('\\t');
    expect(container.textContent).not.toContain('\\nCommute');
  });

  it('renders an accidental confirmed-plan executor prompt as an internal execution card', () => {
    render([
      {
        id: 'u-1',
        role: 'user',
        text:
          'Execute this user-confirmed plan in the open Microsoft 365 surface.\n' +
          '<confirmed_plan>\n' +
          'original_request: /visualize this\n' +
          'intent: visualize\n' +
          'surface: excel\n' +
          'scope: selection\n' +
          'step 1: Insert a new worksheet for the chart\n' +
          'step 2: Create a chart on the new worksheet\n' +
          '</confirmed_plan>',
      },
    ]);

    const card = container.querySelector('.cmd-card-internal');
    expect(card).not.toBeNull();
    expect(card?.textContent).toContain('visualize plan');
    expect(card?.textContent).toContain('excel');
    expect(card?.textContent).toContain('Insert a new worksheet');
    expect(container.textContent).not.toContain('<confirmed_plan>');
    expect(container.textContent).not.toContain('Treat the plan as approved intent');
  });

  it('renders a source excerpt as an inert quotation in the citation peek', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        sources: [
          {
            title: 'Policy',
            locator: '§3.2',
            excerpt: 'Systems must sustain 99.9% <b>availability</b>.',
          },
        ],
      },
    ]);
    openFirstCitation();
    const quote = container.querySelector('blockquote.cite-excerpt');
    expect(quote).not.toBeNull();
    // Untrusted source text renders literally — no markup is interpreted.
    expect(quote?.querySelector('b')).toBeNull();
    expect(quote?.textContent).toContain('Systems must sustain 99.9% <b>availability</b>.');
  });

  it('renders a javascript: citation URI as inert text, never as an executable href', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        // eslint-disable-next-line no-script-url
        sources: [{ title: 'Evil', uri: 'javascript:alert(1)' }],
      },
    ]);
    openFirstCitation();
    // No anchor must be produced for an unsafe scheme.
    expect(container.querySelector('a.cite-detail-link')).toBeNull();
    // The raw URI text is still surfaced inertly inside the detail.
    const detail = container.querySelector('.cite-detail');
    expect(detail?.textContent).toContain('javascript:alert(1)');
  });

  it('renders a data: citation URI as inert text', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        sources: [{ title: 'Sneaky', uri: 'data:text/html,<script>x</script>' }],
      },
    ]);
    openFirstCitation();
    expect(container.querySelector('a.cite-detail-link')).toBeNull();
  });

  it('falls back to a "no link" message when a malformed citation URI cannot be parsed', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        // An opaque string with no scheme and a control char that breaks URL parsing.
        sources: [{ title: 'Broken', uri: 'ht\ntp://bad uri\0' }],
      },
    ]);
    openFirstCitation();
    expect(container.querySelector('a.cite-detail-link')).toBeNull();
    const detail = container.querySelector('.cite-detail');
    // Either the raw uri or the explicit fallback copy — never an anchor.
    expect(detail?.textContent?.length ?? 0).toBeGreaterThan(0);
  });

  it('shows the explicit "No link available" copy when a source has no URI at all', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        sources: [{ title: 'SLA addendum', locator: 'p. 4' }],
      },
    ]);
    openFirstCitation();
    expect(container.querySelector('a.cite-detail-link')).toBeNull();
    expect(container.querySelector('.cite-detail')?.textContent).toContain(
      'No link available for this source.',
    );
  });

  it('toggles the citation popover open and closed and reflects aria-expanded', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: 'grounded',
        sources: [{ title: 'Policy', uri: 'https://example.com' }],
      },
    ]);
    const cite = container.querySelector<HTMLButtonElement>('.cite-btn');
    expect(cite?.getAttribute('aria-expanded')).toBe('false');
    act(() => cite?.click());
    expect(cite?.getAttribute('aria-expanded')).toBe('true');
    expect(container.querySelector('.cite-detail')).not.toBeNull();
    act(() => cite?.click());
    expect(cite?.getAttribute('aria-expanded')).toBe('false');
    expect(container.querySelector('.cite-detail')).toBeNull();
  });

  it('renders a per-message error as an assertive alert', () => {
    render([{ id: 'a-1', role: 'assistant', text: '', error: 'Stream dropped — retry.' }]);
    const alert = container.querySelector('.msg-error[role="alert"]');
    expect(alert).not.toBeNull();
    expect(alert?.textContent).toContain('Stream dropped — retry.');
  });

  it('marks a cancelled turn distinctly and does not render an error', () => {
    render([{ id: 'a-1', role: 'assistant', text: 'partial', cancelled: true }]);
    expect(container.textContent).toContain('Cancelled.');
    expect(container.querySelector('.msg-error')).toBeNull();
  });

  it('renders the streaming caret only on a streaming message', () => {
    render([
      { id: 'a-1', role: 'assistant', text: 'done' },
      { id: 'a-2', role: 'assistant', text: 'typing', streaming: true },
    ]);
    expect(container.querySelectorAll('.caret').length).toBe(1);
  });

  it('renders assistant Markdown headings, emphasis, and tables as structured UI', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: [
          '### Example Time-Blocked Schedule',
          '',
          '**Recommended Action**: Use time blocking.',
          '',
          '| Time Slot | Monday | Tuesday |',
          '|---|---|---|',
          '| 08:00 AM - 10:00 AM | Deep Work | Meetings |',
        ].join('\n'),
      },
    ]);

    expect(container.querySelector('.md-content h4')?.textContent).toBe(
      'Example Time-Blocked Schedule',
    );
    expect(container.querySelector('strong')?.textContent).toBe('Recommended Action');
    const table = container.querySelector<HTMLTableElement>('table.md-table');
    expect(table).not.toBeNull();
    expect(table?.querySelectorAll('th').length).toBe(3);
    expect(table?.textContent).toContain('Deep Work');
    expect(container.textContent).not.toContain('|---|');
  });

  it('exposes assistant Markdown tables as insertable artifacts', () => {
    const onInsertArtifact = vi.fn();
    render(
      [
        {
          id: 'a-1',
          role: 'assistant',
          text: ['| Time | Monday |', '|---|---|', '| 08:00 | Deep Work |'].join('\n'),
        },
      ],
      { onInsertArtifact },
    );

    const insert = container.querySelector<HTMLButtonElement>(
      '.md-table-artifact .md-artifact-insert',
    );
    expect(insert).not.toBeNull();
    expect(insert?.disabled).toBe(false);

    act(() => insert?.click());
    expect(onInsertArtifact).toHaveBeenCalledWith({
      kind: 'markdown-table',
      title: 'Inserted table',
      headers: ['Time', 'Monday'],
      rows: [['08:00', 'Deep Work']],
    });
  });

  it('exposes assistant code blocks as insertable artifacts and respects disabled state', () => {
    const onInsertArtifact = vi.fn();
    render(
      [
        {
          id: 'a-1',
          role: 'assistant',
          text: ['```csv', 'Activity,Hours', 'Deep Work,4', '```'].join('\n'),
        },
      ],
      { onInsertArtifact, insertArtifactDisabledReason: 'Select a destination range first.' },
    );

    const insert = container.querySelector<HTMLButtonElement>(
      '.md-code-artifact .md-artifact-insert',
    );
    expect(insert).not.toBeNull();
    expect(insert?.disabled).toBe(true);
    expect(insert?.title).toBe('Select a destination range first.');

    act(() => insert?.click());
    expect(onInsertArtifact).not.toHaveBeenCalled();
  });

  it('turns inline code host locations into reveal buttons when a reveal callback is available', () => {
    const onRevealLocation = vi.fn();
    render(
      [
        {
          id: 'a-1',
          role: 'assistant',
          text: 'The summary table is in `K6:L18` and the chart starts near `A28`.',
        },
      ],
      { surface: 'excel', onRevealLocation },
    );

    const locations = [...container.querySelectorAll<HTMLButtonElement>('.md-host-location')];
    expect(locations.map((button) => button.textContent)).toEqual(['K6:L18', 'A28']);

    act(() => locations[0]?.click());
    expect(onRevealLocation).toHaveBeenCalledWith('K6:L18');
  });

  it('turns citation: Markdown links into host reveal buttons', () => {
    const onRevealLocation = vi.fn();
    render(
      [
        {
          id: 'a-1',
          role: 'assistant',
          text: "The summary is in [Weekly Hours](<citation:'Daily schedule'!K6:L18>).",
        },
      ],
      { surface: 'excel', onRevealLocation },
    );

    const location = container.querySelector<HTMLButtonElement>('.md-host-location');
    expect(location).not.toBeNull();
    expect(location?.textContent).toBe('Weekly Hours');

    act(() => location?.click());
    expect(onRevealLocation).toHaveBeenCalledWith("citation:'Daily schedule'!K6:L18");
  });

  it('turns Word paragraph citations into host reveal buttons on Word only', () => {
    const onRevealLocation = vi.fn();
    render(
      [{ id: 'a-1', role: 'assistant', text: 'See [the paragraph](<citation:paragraph:7>).' }],
      { surface: 'word', onRevealLocation },
    );

    const location = container.querySelector<HTMLButtonElement>('.md-host-location');
    expect(location).not.toBeNull();
    act(() => location?.click());
    expect(onRevealLocation).toHaveBeenCalledWith('citation:paragraph:7');
  });

  it('does not render Excel-looking inline locations as host links on non-Excel surfaces', () => {
    const onRevealLocation = vi.fn();
    render(
      [
        {
          id: 'a-1',
          role: 'assistant',
          text: 'The summary table is in `K6:L18`.',
        },
      ],
      { surface: 'word', onRevealLocation },
    );

    expect(container.querySelector('.md-host-location')).toBeNull();
    expect(container.querySelector('code')?.textContent).toBe('K6:L18');
  });

  it('does not turn unsafe Markdown links into anchors', () => {
    render([
      {
        id: 'a-1',
        role: 'assistant',
        text: '[bad](javascript:alert(1)) and [good](https://example.com)',
      },
    ]);

    const links = [...container.querySelectorAll<HTMLAnchorElement>('.md-link')];
    expect(links).toHaveLength(1);
    expect(links[0]?.href).toBe('https://example.com/');
    expect(container.textContent).toContain('bad (javascript:alert(1))');
  });

  it('renders raw HTML in assistant text inertly', () => {
    render([{ id: 'a-1', role: 'assistant', text: '<img src=x onerror=alert(1)>' }]);
    expect(container.querySelector('img')).toBeNull();
    expect(container.textContent).toContain('<img src=x onerror=alert(1)>');
  });

  it('renders posted user command, grounding, and scope tokens as chips', () => {
    render([
      {
        id: 'u-1',
        role: 'user',
        text: '/visualize @this create a chart scope:range(A1:B8)',
      },
    ]);

    expect(container.querySelector('.msg-token-command')?.textContent).toBe('/visualize');
    expect(container.querySelector('.msg-token-mention')?.textContent).toBe('@this');
    expect(container.querySelector('.msg-token-scope')?.textContent).toBe('scope:range(A1:B8)');
    expect(container.querySelector('.m.u')?.textContent).toContain('create a chart');
  });

  it('renders user text with token highlighting inertly, never as HTML', () => {
    render([
      {
        id: 'u-1',
        role: 'user',
        text: '@this <img src=x onerror=alert(1)>',
      },
    ]);

    expect(container.querySelector('img')).toBeNull();
    expect(container.querySelector('.msg-token-mention')?.textContent).toBe('@this');
    expect(container.textContent).toContain('<img src=x onerror=alert(1)>');
  });

  it('distinguishes user and assistant bubbles by class', () => {
    render([
      { id: 'u-1', role: 'user', text: 'hi' },
      { id: 'a-1', role: 'assistant', text: 'hello' },
    ]);
    expect(container.querySelector('.m.u')).not.toBeNull();
    expect(container.querySelector('.m.a')).not.toBeNull();
  });

  it('omits the citations block when an assistant message has an empty sources array', () => {
    render([{ id: 'a-1', role: 'assistant', text: 'no sources', sources: [] }]);
    expect(container.querySelector('.cites')).toBeNull();
  });
});
