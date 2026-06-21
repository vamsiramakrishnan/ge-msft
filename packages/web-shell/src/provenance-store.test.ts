import { describe, it, expect } from 'vitest';
import type { ActuationResult, ProvenancePayload } from '@ge/contracts';
import { ProvenanceStore } from './provenance-store.js';

const prov: ProvenancePayload = {
  agentId: 'gemini-enterprise:eng',
  identity: 'v.k@acme',
  timestamp: '2026-06-21T00:00:00.000Z',
  sources: [{ title: 'Vendor Policy', uri: 'https://x' }],
  contentHash: 'sha256:abc',
};

describe('ProvenanceStore', () => {
  it('records a successful change with provenance', () => {
    const store = new ProvenanceStore(() => '2026-06-21T12:00:00.000Z');
    const result: ActuationResult = {
      ok: true,
      changeId: 'c1',
      kind: 'tracked-change',
      location: 'para:3',
    };
    const rec = store.record(result, prov);
    expect(rec).toMatchObject({
      changeId: 'c1',
      kind: 'tracked-change',
      ok: true,
      location: 'para:3',
      at: '2026-06-21T12:00:00.000Z',
    });
    expect(rec.provenance?.identity).toBe('v.k@acme');
    expect(store.get('c1')).toBe(rec);
    expect(store.size).toBe(1);
  });

  it('records a degraded / failed outcome with its error', () => {
    const store = new ProvenanceStore();
    const result: ActuationResult = {
      ok: false,
      changeId: 'c2',
      kind: 'tracked-change',
      degraded: true,
      error: { code: 'anchor_drift', message: 'gone' },
    };
    const rec = store.record(result);
    expect(rec.ok).toBe(false);
    expect(rec.degraded).toBe(true);
    expect(rec.error).toEqual({ code: 'anchor_drift', message: 'gone' });
    expect(rec.provenance).toBeUndefined();
  });

  it('lists records and overwrites by changeId', () => {
    const store = new ProvenanceStore();
    store.record({ ok: true, changeId: 'c1', kind: 'write-cells' });
    store.record({
      ok: false,
      changeId: 'c1',
      kind: 'write-cells',
      error: { code: 'x', message: 'y' },
    });
    expect(store.size).toBe(1);
    expect(store.list()[0]?.ok).toBe(false);
  });
});
