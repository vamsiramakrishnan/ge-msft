import { describe, it, expect } from 'vitest';
import {
  AssistRequestSchema,
  FindingSchema,
  ProvenancePayloadSchema,
  SseEventSchema,
  UnitDescriptorSchema,
  serializeSseEvent,
  parseSseEvent,
  type SseEvent,
} from './index.js';

describe('contracts — CONTRACTS.md example payloads', () => {
  it('parses the POST /review example request', () => {
    // Verbatim from docs/CONTRACTS.md
    const review = {
      intent: 'review',
      unit: {
        notebookId: 'nb_vendor_risk_7f3',
        connectors: [{ type: 'sharepoint', mode: 'federated', scope: 'sites/DealRoom' }],
        restrictToNotebook: false,
        surfaceContext: { kind: 'word', bodyOoxml: '<...>' },
      },
    };
    const parsed = AssistRequestSchema.parse(review);
    expect(parsed.intent).toBe('review');
    expect(parsed.unit.connectors[0]?.mode).toBe('federated');
  });

  it('parses the §7 anchored finding example', () => {
    // Verbatim shape from docs/01-architecture.md §7
    const finding = {
      id: 'f1',
      category: 'policy',
      matchText: '99.5% of the time',
      contextHint: 'Services are available',
      title: 'Availability below FSI standard',
      why: 'Below the FSI availability floor.',
      suggestion: '99.9% of the time',
      sources: [{ title: 'Vendor Risk Policy v4 §3.2', uri: 'https://example/policy' }],
      confidence: 0.91,
      hash: 'a1f9c4e2',
    };
    expect(FindingSchema.parse(finding).confidence).toBeCloseTo(0.91);
  });

  it('parses a well-formed provenance payload', () => {
    const payload = {
      agentId: 'review-agent@v2',
      identity: 'v.k@acme',
      timestamp: new Date().toISOString(),
      sources: [{ title: 'Vendor Risk Policy v4 §3.2' }],
      contentHash: 'deadbeef',
      sessionId: 'sess_123',
    };
    expect(ProvenancePayloadSchema.parse(payload).agentId).toBe('review-agent@v2');
  });

  it('accepts every SurfaceContext kind', () => {
    const kinds = [
      { kind: 'word', selection: 'hi' },
      { kind: 'excel', range: 'A1:B2', values: [['x', 'y']] },
      { kind: 'powerpoint', slideText: 's' },
      { kind: 'onenote', pageId: 'p', sources: ['s'] },
      { kind: 'teams', transcriptWindow: 't' },
    ];
    for (const surfaceContext of kinds) {
      expect(() => UnitDescriptorSchema.parse({ connectors: [], surfaceContext })).not.toThrow();
    }
  });
});

describe('contracts — validation rejects bad shapes', () => {
  it('rejects confidence out of [0,1]', () => {
    const bad = { ...exampleFinding(), confidence: 1.5 };
    expect(() => FindingSchema.parse(bad)).toThrow();
  });

  it('rejects an unknown intent', () => {
    expect(() => AssistRequestSchema.parse({ intent: 'nope', unit: minimalUnit() })).toThrow();
  });

  it('rejects an unknown surface kind', () => {
    expect(() =>
      UnitDescriptorSchema.parse({ connectors: [], surfaceContext: { kind: 'pdf' } }),
    ).toThrow();
  });
});

describe('contracts — SSE wire format', () => {
  const events: SseEvent[] = [
    { type: 'token', text: 'hello ' },
    { type: 'citation', source: { title: 'Policy v4' } },
    {
      type: 'provenance',
      payload: {
        agentId: 'assist@v1',
        identity: 'v.k@acme',
        timestamp: new Date().toISOString(),
        sources: [],
        contentHash: 'abc',
      },
    },
    { type: 'done' },
  ];

  it('round-trips serialize -> parse for each event type', () => {
    for (const ev of events) {
      const wire = serializeSseEvent(ev);
      expect(wire.startsWith(`event: ${ev.type}\n`)).toBe(true);
      expect(wire.endsWith('\n\n')).toBe(true);
      expect(parseSseEvent(wire)).toEqual(ev);
    }
  });

  it('validates the parsed event against the schema', () => {
    const wire = serializeSseEvent({ type: 'token', text: 'x' });
    expect(() => SseEventSchema.parse(parseSseEvent(wire))).not.toThrow();
  });
});

function exampleFinding() {
  return {
    id: 'f1',
    category: 'style' as const,
    matchText: 'foo',
    title: 't',
    why: 'w',
    sources: [],
    confidence: 0.5,
    hash: 'h',
  };
}

function minimalUnit() {
  return { connectors: [], surfaceContext: { kind: 'word' } };
}
