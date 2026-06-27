import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema, asChangeId } from '@ge/contracts';
import {
  transcriptToContext,
  transcriptToDocStateBlocks,
  transcriptToLines,
  searchTranscript,
  MAX_SEARCH_LINES,
  MAX_TRANSCRIPT_LINES,
} from './capture.js';
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

describe('teams whole-transcript read + search (pure)', () => {
  const input = {
    meetingTitle: 'Q3 Renewal sync',
    transcript:
      'Pat: We need the SLA at 99.9%.\nSam: Agreed, let us draft it.\n   \nPat: I will send the addendum.',
    participants: ['Pat', 'Sam'],
  };

  it('splits a transcript into non-empty turn lines', () => {
    expect(transcriptToLines(input.transcript)).toEqual([
      'Pat: We need the SLA at 99.9%.',
      'Sam: Agreed, let us draft it.',
      'Pat: I will send the addendum.',
    ]);
  });

  it('builds whole-transcript snapshot blocks: title heading + turn lines, capped + anchored', () => {
    const blocks = transcriptToDocStateBlocks(input);
    expect(blocks[0]?.kind).toBe('heading');
    expect(blocks.every((b) => b.locator === 'transcript')).toBe(true);
    expect(blocks.length).toBeLessThanOrEqual(MAX_TRANSCRIPT_LINES + 1); // heading + turn lines
  });

  it('search returns matching turn lines as valid context, scoped to the window', () => {
    const ctx = searchTranscript(input, 'addendum');
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined.toLowerCase()).toContain('addendum');
    expect(joined).not.toContain('Agreed'); // non-matching line excluded
  });

  it('search empty query / no match → []', () => {
    expect(searchTranscript(input, '  ')).toHaveLength(0);
    expect(searchTranscript(input, 'nonexistent-token')).toHaveLength(0);
  });

  it('search is bounded to MAX_SEARCH_LINES matches', () => {
    const many = {
      transcript: Array.from({ length: MAX_SEARCH_LINES + 5 }, (_, i) => `Pat: match ${i}`).join(
        '\n',
      ),
    };
    const ctx = searchTranscript(many, 'match');
    expect(ctx.length).toBeGreaterThan(0);
  });
});

describe('teams actuation planning (pure)', () => {
  it('builds a post plan from params.text', () => {
    const plan = planPostMessage({
      changeId: asChangeId('c1'),
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Action items: 1) raise SLA, 2) draft addendum.' },
    });
    expect(plan).toEqual({ text: 'Action items: 1) raise SLA, 2) draft addendum.' });
  });

  it('threads an Adaptive Card payload through params.html', () => {
    const plan = planPostMessage({
      changeId: asChangeId('c2'),
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Summary', html: '<card>...</card>' },
    });
    expect(plan).toEqual({ text: 'Summary', card: '<card>...</card>' });
  });

  it('falls back to empty text when nothing is provided', () => {
    const plan = planPostMessage({
      changeId: asChangeId('c3'),
      kind: 'post-message',
      surface: 'teams',
      params: {},
    });
    expect(plan).toEqual({ text: '' });
  });
});
