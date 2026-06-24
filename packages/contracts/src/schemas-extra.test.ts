import { describe, it, expect } from 'vitest';
import {
  asChangeId,
  asSessionId,
  ChangeIdSchema,
  SessionIdSchema,
  FindingSchema,
  SourceRefSchema,
  ConnectorRefSchema,
  SurfaceContextSchema,
  UnitDescriptorSchema,
  ProvenancePayloadSchema,
  AssistRequestSchema,
  ActionRequestSchema,
  IntentSchema,
} from './index.js';

/**
 * Behavioral coverage for the contract schemas' rejection paths and the two id brand
 * mint points. The existing `contracts.test.ts` covers the happy-path example payloads;
 * this file targets the failure branches an untrusted payload would hit.
 */

describe('brand — id mint points', () => {
  it('asChangeId returns the same underlying string value (brand erased at runtime)', () => {
    const id = asChangeId('chg_abc');
    expect(id).toBe('chg_abc');
    // The branded value is still parseable by its own schema.
    expect(ChangeIdSchema.parse(id)).toBe('chg_abc');
  });

  it('asSessionId returns the same underlying string value', () => {
    const id = asSessionId('sess_123');
    expect(id).toBe('sess_123');
    expect(SessionIdSchema.parse(id)).toBe('sess_123');
  });

  it('asChangeId rejects a non-string at the boundary', () => {
    expect(() => asChangeId(42 as unknown as string)).toThrow();
  });

  it('asSessionId rejects a non-string at the boundary', () => {
    expect(() => asSessionId(null as unknown as string)).toThrow();
  });

  it('an empty string is still a valid id (brand adds no length constraint)', () => {
    expect(asChangeId('')).toBe('');
    expect(asSessionId('')).toBe('');
  });
});

describe('SourceRef schema', () => {
  it('accepts a title-only source', () => {
    expect(SourceRefSchema.parse({ title: 'Policy v4' })).toEqual({ title: 'Policy v4' });
  });

  it('rejects a source with no title', () => {
    expect(() => SourceRefSchema.parse({ uri: 'https://x' })).toThrow();
  });

  it('rejects a non-string uri', () => {
    expect(() => SourceRefSchema.parse({ title: 't', uri: 5 })).toThrow();
  });
});

describe('Finding schema — rejection paths', () => {
  const good = {
    id: 'f1',
    category: 'policy' as const,
    matchText: 'foo',
    title: 't',
    why: 'w',
    sources: [],
    confidence: 0.5,
    hash: 'h',
  };

  it('accepts a minimal valid finding (no optional suggestion/contextHint)', () => {
    expect(FindingSchema.parse(good).suggestion).toBeUndefined();
  });

  it('rejects an unknown category', () => {
    expect(() => FindingSchema.parse({ ...good, category: 'grammar' })).toThrow();
  });

  it('rejects confidence below 0', () => {
    expect(() => FindingSchema.parse({ ...good, confidence: -0.01 })).toThrow();
  });

  it('rejects confidence above 1', () => {
    expect(() => FindingSchema.parse({ ...good, confidence: 1.0001 })).toThrow();
  });

  it('accepts confidence at the 0 and 1 boundaries', () => {
    expect(FindingSchema.parse({ ...good, confidence: 0 }).confidence).toBe(0);
    expect(FindingSchema.parse({ ...good, confidence: 1 }).confidence).toBe(1);
  });

  it('rejects a missing hash (provenance is mandatory)', () => {
    const { hash: _hash, ...noHash } = good;
    expect(() => FindingSchema.parse(noHash)).toThrow();
  });

  it('rejects a missing matchText (the anchor is mandatory)', () => {
    const { matchText: _m, ...noMatch } = good;
    expect(() => FindingSchema.parse(noMatch)).toThrow();
  });

  it('rejects a source entry that lacks a title', () => {
    expect(() => FindingSchema.parse({ ...good, sources: [{ uri: 'https://x' }] })).toThrow();
  });
});

describe('ConnectorRef schema', () => {
  it('accepts a federated sharepoint connector with scope', () => {
    const c = ConnectorRefSchema.parse({
      type: 'sharepoint',
      mode: 'federated',
      scope: 'sites/DealRoom',
    });
    expect(c.scope).toBe('sites/DealRoom');
  });

  it('rejects an unknown connector type', () => {
    expect(() => ConnectorRefSchema.parse({ type: 'dropbox', mode: 'federated' })).toThrow();
  });

  it('rejects an unknown mode', () => {
    expect(() => ConnectorRefSchema.parse({ type: 'onedrive', mode: 'streaming' })).toThrow();
  });

  it('rejects a missing mode', () => {
    expect(() => ConnectorRefSchema.parse({ type: 'onedrive' })).toThrow();
  });
});

describe('SurfaceContext discriminated union', () => {
  it('rejects an unknown kind', () => {
    expect(() => SurfaceContextSchema.parse({ kind: 'pdf' })).toThrow();
  });

  it('rejects a missing discriminator', () => {
    expect(() => SurfaceContextSchema.parse({ selection: 'hi' })).toThrow();
  });

  it('rejects excel values that are not a 2D string array', () => {
    expect(() =>
      SurfaceContextSchema.parse({ kind: 'excel', range: 'A1', values: [[1, 2]] }),
    ).toThrow();
  });

  it('accepts the outlook surface (mail capture)', () => {
    const ctx = SurfaceContextSchema.parse({
      kind: 'outlook',
      subject: 's',
      body: 'b',
      from: 'a@b',
    });
    expect(ctx.kind).toBe('outlook');
  });

  it('rejects onenote sources that are not strings', () => {
    expect(() =>
      SurfaceContextSchema.parse({ kind: 'onenote', sources: [{ not: 'a string' }] }),
    ).toThrow();
  });
});

describe('UnitDescriptor schema', () => {
  it('rejects a missing surfaceContext', () => {
    expect(() => UnitDescriptorSchema.parse({ connectors: [] })).toThrow();
  });

  it('rejects connectors that is not an array', () => {
    expect(() =>
      UnitDescriptorSchema.parse({ connectors: {}, surfaceContext: { kind: 'word' } }),
    ).toThrow();
  });

  it('rejects a non-boolean restrictToNotebook', () => {
    expect(() =>
      UnitDescriptorSchema.parse({
        connectors: [],
        restrictToNotebook: 'yes',
        surfaceContext: { kind: 'word' },
      }),
    ).toThrow();
  });

  it('propagates a bad connector inside the array', () => {
    expect(() =>
      UnitDescriptorSchema.parse({
        connectors: [{ type: 'sharepoint', mode: 'bogus' }],
        surfaceContext: { kind: 'word' },
      }),
    ).toThrow();
  });
});

describe('ProvenancePayload schema', () => {
  const good = {
    agentId: 'review-agent@v2',
    identity: 'v.k@acme',
    timestamp: new Date().toISOString(),
    sources: [],
    contentHash: 'deadbeef',
  };

  it('accepts a payload without the optional sessionId', () => {
    expect(ProvenancePayloadSchema.parse(good).sessionId).toBeUndefined();
  });

  it('accepts a branded sessionId', () => {
    const parsed = ProvenancePayloadSchema.parse({ ...good, sessionId: 'sess_9' });
    expect(parsed.sessionId).toBe('sess_9');
  });

  it('rejects a missing contentHash (writes must be hashed for reversibility)', () => {
    const { contentHash: _c, ...noHash } = good;
    expect(() => ProvenancePayloadSchema.parse(noHash)).toThrow();
  });

  it('rejects a missing identity (writes are scoped to the signed-in user)', () => {
    const { identity: _i, ...noId } = good;
    expect(() => ProvenancePayloadSchema.parse(noId)).toThrow();
  });

  it('rejects a non-string sessionId', () => {
    expect(() => ProvenancePayloadSchema.parse({ ...good, sessionId: 99 })).toThrow();
  });
});

describe('Intent schema', () => {
  it('accepts every declared verb', () => {
    for (const verb of ['ask', 'summarize', 'explain', 'rewrite', 'review', 'draft', 'notes']) {
      expect(IntentSchema.parse(verb)).toBe(verb);
    }
  });

  it('rejects an undeclared verb', () => {
    expect(() => IntentSchema.parse('delete-everything')).toThrow();
  });
});

describe('AssistRequest schema', () => {
  const unit = { connectors: [], surfaceContext: { kind: 'word' as const } };

  it('accepts an assist request with a query and target', () => {
    const r = AssistRequestSchema.parse({
      intent: 'ask',
      unit,
      query: 'summarize',
      target: { range: 'A1' },
      changeId: 'chg_1',
    });
    expect(r.target?.range).toBe('A1');
  });

  it('rejects a request missing the unit', () => {
    expect(() => AssistRequestSchema.parse({ intent: 'ask' })).toThrow();
  });

  it('rejects a request whose unit is malformed', () => {
    expect(() => AssistRequestSchema.parse({ intent: 'ask', unit: { connectors: [] } })).toThrow();
  });

  it('rejects a non-string query', () => {
    expect(() => AssistRequestSchema.parse({ intent: 'ask', unit, query: 123 })).toThrow();
  });
});

describe('ActionRequest schema', () => {
  it('accepts a checkout action without a payload', () => {
    const r = ActionRequestSchema.parse({
      action: 'checkout',
      connector: 'sharepoint',
      target: 'sites/X/file.docx',
      changeId: 'chg_2',
    });
    expect(r.payload).toBeUndefined();
  });

  it('accepts an upload action with a base64 payload', () => {
    const r = ActionRequestSchema.parse({
      action: 'upload',
      connector: 'onedrive',
      target: 'path/x',
      payload: { filename: 'x.docx', contentBase64: 'AAAA' },
      changeId: 'chg_3',
    });
    expect(r.payload?.filename).toBe('x.docx');
  });

  it('rejects an unknown action', () => {
    expect(() =>
      ActionRequestSchema.parse({
        action: 'purge',
        connector: 'onedrive',
        target: 't',
        changeId: 'c',
      }),
    ).toThrow();
  });

  it('rejects a missing changeId (actions must be idempotent)', () => {
    expect(() =>
      ActionRequestSchema.parse({ action: 'download', connector: 'onedrive', target: 't' }),
    ).toThrow();
  });

  it('rejects a payload missing contentBase64', () => {
    expect(() =>
      ActionRequestSchema.parse({
        action: 'upload',
        connector: 'onedrive',
        target: 't',
        payload: { filename: 'x' },
        changeId: 'c',
      }),
    ).toThrow();
  });
});
