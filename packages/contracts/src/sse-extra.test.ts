import { describe, it, expect } from 'vitest';
import { SseEventSchema, serializeSseEvent, parseSseEvent, type SseEvent } from './index.js';

/**
 * Behavioral coverage for the SSE wire protocol: per-event-type round-trips, the
 * parse failure paths (no data line, malformed JSON, schema-invalid payloads), and
 * the `data:`-with/without-space and multiline accumulation branches in parseSseEvent.
 */

describe('SSE — every event variant round-trips', () => {
  const variants: SseEvent[] = [
    { type: 'token', text: 'hi' },
    { type: 'activity', text: 'Analyzing selected range' },
    {
      type: 'finding',
      finding: {
        id: 'f1',
        category: 'style',
        matchText: 'm',
        title: 't',
        why: 'w',
        sources: [],
        confidence: 0.5,
        hash: 'h',
      },
    },
    {
      type: 'slide',
      title: 'Intro',
      bullets: ['a', 'b'],
      sources: [{ title: 'src' }],
    },
    { type: 'citation', source: { title: 'Policy', uri: 'https://x' } },
    { type: 'code-execution', language: 'python', code: 'print(1)' },
    { type: 'code-execution-result', outcome: 'OUTCOME_OK', output: '1\n' },
    {
      type: 'grounding-support',
      start: 0,
      end: 5,
      score: 0.9,
      sources: [{ title: 's' }],
    },
    { type: 'policy', verdict: 'block', reason: 'banned phrase' },
    { type: 'related-questions', questions: ['why?', 'how?'] },
    {
      type: 'provenance',
      payload: {
        agentId: 'a@v1',
        identity: 'u@acme',
        timestamp: new Date().toISOString(),
        sources: [],
        contentHash: 'abc',
      },
    },
    { type: 'error', code: 'E_X', message: 'boom' },
    { type: 'done' },
  ];

  it('serialize -> parse is identity for each variant', () => {
    for (const ev of variants) {
      const wire = serializeSseEvent(ev);
      expect(wire.startsWith(`event: ${ev.type}\n`)).toBe(true);
      expect(wire.endsWith('\n\n')).toBe(true);
      expect(parseSseEvent(wire)).toEqual(ev);
    }
  });
});

describe('SSE — schema rejection paths', () => {
  it('rejects an unknown event type', () => {
    expect(() => SseEventSchema.parse({ type: 'heartbeat' })).toThrow();
  });

  it('rejects grounding-support with a negative index', () => {
    expect(() =>
      SseEventSchema.parse({ type: 'grounding-support', start: -1, end: 2, sources: [] }),
    ).toThrow();
  });

  it('rejects grounding-support with a non-integer index', () => {
    expect(() =>
      SseEventSchema.parse({ type: 'grounding-support', start: 1.5, end: 2, sources: [] }),
    ).toThrow();
  });

  it('accepts grounding-support without the optional score', () => {
    const ev = SseEventSchema.parse({
      type: 'grounding-support',
      start: 0,
      end: 1,
      sources: [{ title: 's' }],
    });
    expect(ev).toMatchObject({ type: 'grounding-support', start: 0, end: 1 });
  });

  it('rejects a policy verdict outside the allowed enum', () => {
    expect(() => SseEventSchema.parse({ type: 'policy', verdict: 'maybe' })).toThrow();
  });

  it('rejects a token event with a non-string text', () => {
    expect(() => SseEventSchema.parse({ type: 'token', text: 7 })).toThrow();
  });

  it('rejects code-execution for any language other than python', () => {
    expect(() =>
      SseEventSchema.parse({ type: 'code-execution', language: 'javascript', code: 'alert(1)' }),
    ).toThrow();
  });

  it('rejects code-execution-result with an unknown outcome', () => {
    expect(() =>
      SseEventSchema.parse({ type: 'code-execution-result', outcome: 'OUTCOME_SIDE_EFFECT' }),
    ).toThrow();
  });

  it('rejects an error event missing the message', () => {
    expect(() => SseEventSchema.parse({ type: 'error', code: 'E' })).toThrow();
  });

  it('rejects a finding event carrying an invalid finding', () => {
    expect(() => SseEventSchema.parse({ type: 'finding', finding: { id: 'x' } })).toThrow();
  });
});

describe('parseSseEvent — block parsing branches', () => {
  it('throws a clear error when the block has no data line', () => {
    expect(() => parseSseEvent('event: token\n')).toThrow(/no data line/);
  });

  it('throws when there are only comment/other lines (no data:)', () => {
    expect(() => parseSseEvent(': keep-alive\nevent: token\n')).toThrow(/no data line/);
  });

  it('parses a data line without the conventional space after the colon', () => {
    // `data:{...}` (no space) must slice 5 chars, not 6.
    const ev = parseSseEvent('event: token\ndata:{"type":"token","text":"x"}\n\n');
    expect(ev).toEqual({ type: 'token', text: 'x' });
  });

  it('accumulates a multi-line data payload across several data: lines', () => {
    // Two data: lines that concatenate into one JSON document.
    const block = 'event: token\n' + 'data: {"type":"token",\n' + 'data: "text":"joined"}\n\n';
    expect(parseSseEvent(block)).toEqual({ type: 'token', text: 'joined' });
  });

  it('throws when the data line is not valid JSON', () => {
    expect(() => parseSseEvent('data: {not json}')).toThrow();
  });

  it('throws when the JSON is valid but violates the event schema', () => {
    expect(() => parseSseEvent('data: {"type":"token"}')).toThrow();
  });

  it('tolerates trailing whitespace on the data line (trimEnd)', () => {
    const ev = parseSseEvent('data: {"type":"done"}   \n\n');
    expect(ev).toEqual({ type: 'done' });
  });
});
