import { describe, it, expect } from 'vitest';
import type { ActuationResult, ProvenancePayload } from '@ge/contracts';
import { asChangeId, assessActuationResult } from '@ge/contracts';
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
      changeId: asChangeId('c1'),
      kind: 'tracked-change',
      location: 'para:3',
    };
    const rec = store.record(result, prov);
    expect(rec).toMatchObject({
      changeId: asChangeId('c1'),
      kind: 'tracked-change',
      ok: true,
      location: 'para:3',
      at: '2026-06-21T12:00:00.000Z',
    });
    expect(rec.provenance?.identity).toBe('v.k@acme');
    expect(store.get(asChangeId('c1'))).toEqual(rec);
    expect(store.size).toBe(1);
  });

  it('records a degraded / failed outcome with its error', () => {
    const store = new ProvenanceStore();
    const result: ActuationResult = {
      ok: false,
      changeId: asChangeId('c2'),
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
    store.record({ ok: true, changeId: asChangeId('c1'), kind: 'write-cells' });
    store.record({
      ok: false,
      changeId: asChangeId('c1'),
      kind: 'write-cells',
      error: { code: 'x', message: 'y' },
    });
    expect(store.size).toBe(1);
    expect(store.list()[0]?.ok).toBe(false);
  });

  it('preserves uncertain recovery truth even when the host reported a successful write', () => {
    const store = new ProvenanceStore();
    const result: ActuationResult = {
      kind: 'write-cells',
      changeId: asChangeId('uncertain'),
      ok: true,
      verification: { status: 'verified', afterHash: 'sha256:123' },
      recoveryPending: true,
    };
    const record = store.record(result);
    expect(record.ok).toBe(true);
    expect(record.recoveryPending).toBe(true);
    expect(assessActuationResult(record)).toBe('uncertain');
    result.verification!.afterHash = 'changed';
    result.recoveryPending = false;
    expect(record.verification?.afterHash).toBe('sha256:123');
    expect(assessActuationResult(record)).toBe('uncertain');
  });

  it('keeps unknown verification and error facts detached from later caller mutations', () => {
    const store = new ProvenanceStore();
    const result: ActuationResult = {
      kind: 'write-cells',
      changeId: asChangeId('unknown'),
      ok: false,
      verification: { status: 'unknown' },
      error: { code: 'outcome_unknown', message: 'Inspect cells' },
    };
    const record = store.record(result);
    result.verification!.status = 'verified';
    result.error!.code = 'failed';
    expect(record.verification?.status).toBe('unknown');
    expect(record.error?.code).toBe('outcome_unknown');
    expect(assessActuationResult(record)).toBe('uncertain');
  });

  it('returns detached audit snapshots so consumers cannot rewrite stored verification or attribution', () => {
    const store = new ProvenanceStore();
    const provenance = structuredClone(prov);
    const id = asChangeId('immutable');
    const recorded = store.record(
      {
        kind: 'write-cells',
        changeId: id,
        ok: true,
        verification: { status: 'unknown' },
        inverse: { op: 'restore-values', range: 'A1', values: [['Before']] },
      },
      provenance,
    );
    provenance.sources[0]!.title = 'Changed input';
    recorded.verification!.status = 'verified';
    recorded.provenance!.sources[0]!.title = 'Changed return';
    store.get(id)!.verification!.status = 'verified';
    store.list()[0]!.inverse = undefined;
    expect(assessActuationResult(store.get(id)!)).toBe('uncertain');
    expect(store.get(id)?.provenance?.sources[0]?.title).toBe('Vendor Policy');
    expect(store.get(id)?.inverse).toEqual({
      op: 'restore-values',
      range: 'A1',
      values: [['Before']],
    });
    const shared = store.recordShare(
      { name: 'Handoff', bytes: 5, sourceLabel: 'selection', truncated: false },
      prov,
    );
    shared.provenance!.sources[0]!.title = 'Changed share';
    store.listShares()[0]!.name = 'Changed list';
    expect(store.listShares()[0]).toMatchObject({
      name: 'Handoff',
      provenance: { sources: [{ title: 'Vendor Policy' }] },
    });
  });
});

it('preserves a detached host inverse receipt for a later explicit undo operation', () => {
  const store = new ProvenanceStore();
  const inverse = { op: 'restore-values' as const, range: 'A1', values: [['Before']] };
  const record = store.record({
    ok: true,
    changeId: asChangeId('inverse'),
    kind: 'write-cells',
    inverse,
  });
  inverse.values[0]![0] = 'Mutated after recording';
  expect(record.inverse).toEqual({ op: 'restore-values', range: 'A1', values: [['Before']] });
});
