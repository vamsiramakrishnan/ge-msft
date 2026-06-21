import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema } from '@ge/contracts';
import { transcriptToContext } from './capture.js';
import { planPostMessage } from './actuate-plan.js';

describe('teams capture (pure)', () => {
  it('produces valid text context from a transcript window', () => {
    const ctx = transcriptToContext({
      meetingTitle: 'Q3 Renewal sync',
      transcript: 'Pat: We need the SLA at 99.9%.\nSam: Agreed, let us draft it.',
      participants: ['Pat', 'Sam'],
    });
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.every((c) => c.value.as === 'text')).toBe(true);
  });

  it('labels the source with meeting title and participants', () => {
    const ctx = transcriptToContext({
      meetingTitle: 'Renewal',
      transcript: 'Please confirm the renewal terms.',
      participants: ['Sam', 'Pat'],
    });
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('Renewal');
    expect(joined).toContain('Sam, Pat');
    expect(joined).toContain('renewal terms');
  });

  it('works without metadata (transcript only)', () => {
    const ctx = transcriptToContext({ transcript: 'We are blocked on the contract.' });
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('blocked');
    expect(joined).not.toContain('Meeting:');
  });

  it('tags every chunk to the teams surface', () => {
    const ctx = transcriptToContext({ transcript: 'A short turn.' });
    expect(ctx.every((c) => c.ref.surface === 'teams')).toBe(true);
  });
});

describe('teams actuation planning (pure)', () => {
  it('builds a post plan from params.text', () => {
    const plan = planPostMessage({
      changeId: 'c1',
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Action items: 1) raise SLA, 2) draft addendum.' },
    });
    expect(plan).toEqual({ text: 'Action items: 1) raise SLA, 2) draft addendum.' });
  });

  it('threads an Adaptive Card payload through params.html', () => {
    const plan = planPostMessage({
      changeId: 'c2',
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Summary', html: '<card>...</card>' },
    });
    expect(plan).toEqual({ text: 'Summary', card: '<card>...</card>' });
  });

  it('falls back to empty text when nothing is provided', () => {
    const plan = planPostMessage({
      changeId: 'c3',
      kind: 'post-message',
      surface: 'teams',
      params: {},
    });
    expect(plan).toEqual({ text: '' });
  });
});
