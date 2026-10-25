import { describe, expect, it } from 'vitest';
import { answerForClipboard } from './answer-copy.js';

const answer = [
  '## Decision brief',
  'Approve the **next pilot** phase.',
  '',
  '- Owner: *Rina*',
  '- See [policy](https://example.com/p)',
  '',
  '| Check | Date |',
  '|---|---|',
  '| Gateway | 16 Oct |',
].join('\n');

describe('answerForClipboard', () => {
  it('pastes as plain text without markdown symbols', () => {
    expect(answerForClipboard(answer).plain).toBe(
      [
        'Decision brief',
        'Approve the next pilot phase.',
        '',
        '• Owner: Rina',
        '• See policy (https://example.com/p)',
        '',
        'Check\tDate',
        'Gateway\t16 Oct',
      ].join('\n'),
    );
  });

  it('pastes as formatted HTML for Word and Outlook', () => {
    const { html } = answerForClipboard(answer);
    expect(html).toContain('<h3>Decision brief</h3>');
    expect(html).toContain('<strong>next pilot</strong>');
    expect(html).toContain('<ul><li>Owner: <em>Rina</em></li>');
    expect(html).toContain('<a href="https://example.com/p">policy</a>');
    expect(html).toContain('<th>Check</th>');
    expect(html).toContain('<td>16 Oct</td>');
  });

  it('escapes model text and drops non-http links', () => {
    const { html } = answerForClipboard('<img src=x onerror=alert(1)> [x](javascript:alert(1))');
    expect(html).not.toContain('<img');
    expect(html).not.toContain('javascript:');
    expect(html).toContain('&lt;img');
  });
});
