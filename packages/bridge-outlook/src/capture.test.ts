import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema, asChangeId } from '@ge/contracts';
import { mailItemToContext } from './capture.js';
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
