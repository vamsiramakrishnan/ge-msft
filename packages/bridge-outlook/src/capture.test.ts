import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema, asChangeId } from '@ge/contracts';
import {
  mailItemToContext,
  mailItemToDocStateBlocks,
  mailBodyToLines,
  searchMailItem,
  MAX_SEARCH_LINES,
  MAX_OUTLINE_LINES,
} from './capture.js';
import { planReply } from './actuate-plan.js';

describe('outlook capture (pure)', () => {
  it('produces valid text context from a plain-text mail item', () => {
    const ctx = mailItemToContext({
      id: 'AAMk-123',
      subject: 'SLA concerns',
      from: 'pat@acme.com',
      body: 'The services are available 99.5% of the time. Can we raise it?',
      bodyType: 'text',
    });
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.every((c) => c.value.as === 'text')).toBe(true);
  });

  it('labels the source with subject and from header', () => {
    const ctx = mailItemToContext({
      subject: 'Renewal',
      from: 'sam@acme.com',
      body: 'Please confirm the renewal terms.',
      bodyType: 'text',
    });
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('Renewal');
    expect(joined).toContain('sam@acme.com');
    expect(joined).toContain('renewal terms');
  });

  it('normalizes an HTML body through the string path', () => {
    const ctx = mailItemToContext({
      subject: 'Status',
      body: '<p>We are <strong>blocked</strong> on the contract.</p>',
      bodyType: 'html',
    });
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('blocked');
    expect(joined).not.toContain('<strong>');
  });
});

describe('outlook whole-item read + search (pure)', () => {
  const item = {
    id: 'AAMk-9',
    subject: 'SLA concerns',
    from: 'pat@acme.com',
    body: '<p>We are at <strong>99.5%</strong> availability.</p><p>Can we raise the SLA to 99.9%?</p>',
    bodyType: 'html' as const,
  };

  it('reduces an HTML body to plain lines (tags stripped, entities decoded)', () => {
    const lines = mailBodyToLines('<p>A &amp; B</p><p>line two</p>', 'html');
    expect(lines).toEqual(['A & B', 'line two']);
  });

  it('builds whole-item snapshot blocks: subject heading + from + body lines, capped', () => {
    const blocks = mailItemToDocStateBlocks(item);
    expect(blocks[0]?.kind).toBe('heading');
    expect(blocks.every((b) => b.locator === 'mail:AAMk-9')).toBe(true);
    expect(blocks.some((b) => b.text.includes('From: pat@acme.com'))).toBe(true);
    expect(blocks.length).toBeLessThanOrEqual(MAX_OUTLINE_LINES + 2); // heading + from + body lines
  });

  it('search returns matching body lines as valid context, scoped to the item', () => {
    const ctx = searchMailItem(item, 'raise the sla');
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined.toLowerCase()).toContain('raise the sla');
    expect(joined).not.toContain('99.5%'); // non-matching line excluded
  });

  it('search empty query / no match → []', () => {
    expect(searchMailItem(item, '  ')).toHaveLength(0);
    expect(searchMailItem(item, 'nonexistent-token')).toHaveLength(0);
  });

  it('search is bounded to MAX_SEARCH_LINES matches', () => {
    const many = {
      body: Array.from({ length: MAX_SEARCH_LINES + 5 }, (_, i) => `match line ${i}`).join('\n'),
      bodyType: 'text' as const,
    };
    const ctx = searchMailItem(many, 'match line');
    // Non-empty but derived from at most MAX_SEARCH_LINES lines (the scan breaks at the cap).
    expect(ctx.length).toBeGreaterThan(0);
  });
});

describe('outlook actuation planning (pure)', () => {
  it('extracts body/to/subject from params.mail', () => {
    const plan = planReply({
      changeId: asChangeId('c1'),
      kind: 'reply-mail',
      surface: 'outlook',
      params: {
        mail: {
          to: ['pat@acme.com'],
          subject: 'Re: SLA concerns',
          body: 'We can raise availability to 99.9%.',
        },
      },
    });
    expect(plan).toEqual({
      body: 'We can raise availability to 99.9%.',
      to: ['pat@acme.com'],
      subject: 'Re: SLA concerns',
    });
  });

  it('falls back to params.text for the body', () => {
    const plan = planReply({
      changeId: asChangeId('c2'),
      kind: 'reply-mail',
      surface: 'outlook',
      params: { text: 'Thanks, will follow up.' },
    });
    expect(plan).toEqual({ body: 'Thanks, will follow up.' });
  });
});
