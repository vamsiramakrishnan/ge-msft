import { describe, it, expect } from 'vitest';
import {
  ActuationRequestSchema,
  CapabilityManifestSchema,
  ContextRefSchema,
  ResolvedContextSchema,
} from './index.js';

describe('context + capability model', () => {
  it('parses a context ref and its resolved text value', () => {
    const ref = ContextRefSchema.parse({
      id: 'word:selection',
      kind: 'selection',
      surface: 'word',
      title: 'Selection',
      preview: 'available 99.5% of the time',
      live: true,
    });
    expect(ref.live).toBe(true);
    const resolved = ResolvedContextSchema.parse({
      ref,
      value: { as: 'text', text: 'available 99.5% of the time' },
    });
    expect(resolved.value.as).toBe('text');
  });

  it('parses each resolved-context value kind', () => {
    for (const value of [
      { as: 'text', text: 't' },
      { as: 'indexed-document', documentName: 'projects/x/d' },
      { as: 'drive-document', driveId: 'gd' },
      { as: 'person', displayName: 'V' },
    ]) {
      const base = { id: 'x', kind: 'document', surface: 'word', title: 'X' };
      expect(() => ResolvedContextSchema.parse({ ref: base, value })).not.toThrow();
    }
  });

  it('validates an actuation request with provenance + idempotent changeId', () => {
    const req = ActuationRequestSchema.parse({
      changeId: 'c1',
      kind: 'tracked-change',
      surface: 'word',
      params: {
        text: '99.9% of the time',
        target: { matchText: '99.5% of the time', contextHint: 'Services are available' },
        sources: [{ title: 'Vendor Risk Policy v4 §3.2' }],
      },
      provenance: {
        agentId: 'review@v2',
        identity: 'v.k@acme',
        timestamp: new Date().toISOString(),
        sources: [{ title: 'Vendor Risk Policy v4 §3.2' }],
        contentHash: 'abcd1234',
      },
    });
    expect(req.params.target?.matchText).toBe('99.5% of the time');
  });

  it('parses a surface capability manifest', () => {
    const manifest = CapabilityManifestSchema.parse({
      surface: 'word',
      contextKinds: ['selection', 'document', 'paragraph', 'comment'],
      actuations: [
        {
          kind: 'tracked-change',
          surface: 'word',
          title: 'Insert as tracked change',
          reversible: true,
        },
        { kind: 'comment-reply', surface: 'word', title: 'Reply & resolve', reversible: true },
      ],
    });
    expect(manifest.actuations).toHaveLength(2);
  });

  it('rejects an unknown actuation kind', () => {
    expect(() =>
      ActuationRequestSchema.parse({
        changeId: 'c',
        kind: 'delete-everything',
        surface: 'word',
        params: {},
      }),
    ).toThrow();
  });
});
