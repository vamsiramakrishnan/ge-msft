import { describe, it, expect } from 'vitest';
import { ProvenancePayloadSchema, type ProvenancePayload } from '@ge/contracts';
import { provenanceRecord, provenanceKey } from './provenance-record.js';

const payload: ProvenancePayload = {
  agentId: 'review-agent@v2',
  identity: 'v.k@acme',
  timestamp: '2026-06-22T12:00:00Z',
  contentHash: 'sha256:abc123',
  sources: [{ title: 'SLA Policy', uri: 'https://acme/sla' }],
};

describe('provenanceRecord (Word durable provenance)', () => {
  it('keys the record stably as ge:prov:<changeId>', () => {
    expect(provenanceKey('chg-42')).toBe('ge:prov:chg-42');
    expect(provenanceRecord('chg-42', payload).key).toBe('ge:prov:chg-42');
  });

  it('round-trips the JSON back to the payload (plus the changeId)', () => {
    const { json } = provenanceRecord('chg-7', payload);
    const parsed = JSON.parse(json) as { changeId: string } & ProvenancePayload;
    expect(parsed.changeId).toBe('chg-7');
    const { changeId: _changeId, ...rest } = parsed;
    expect(ProvenancePayloadSchema.parse(rest)).toEqual(payload);
  });

  it('emits a well-formed XML custom-part carrying the key and fields', () => {
    const { xml } = provenanceRecord('chg-9', payload);
    expect(xml.startsWith('<geProvenance')).toBe(true);
    expect(xml.endsWith('</geProvenance>')).toBe(true);
    expect(xml).toContain('key="ge:prov:chg-9"');
    expect(xml).toContain('agentId="review-agent@v2"');
    expect(xml).toContain('<source title="SLA Policy" uri="https://acme/sla"/>');
  });

  it('escapes XML metacharacters so a crafted source cannot break the part', () => {
    const { xml } = provenanceRecord('c', {
      ...payload,
      sources: [{ title: 'A & B <x>' }],
    });
    expect(xml).toContain('title="A &amp; B &lt;x&gt;"');
    expect(xml).not.toContain('<x>');
  });
});
