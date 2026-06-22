import { describe, it, expect } from 'vitest';
import { ProvenancePayloadSchema, asSessionId, type ProvenancePayload } from '@ge/contracts';
import { provenanceRecord, provenanceKey } from './provenance-record.js';

const payload: ProvenancePayload = {
  agentId: 'review-agent@v2',
  identity: 'v.k@acme',
  timestamp: '2026-06-22T12:00:00Z',
  contentHash: 'sha256:abc123',
  sources: [
    { title: 'SLA Policy', uri: 'https://acme/sla', locator: 'p.4' },
    { title: 'Uptime Memo' },
  ],
};

describe('provenanceRecord (Excel durable provenance)', () => {
  it('keys the record stably as ge:prov:<changeId>', () => {
    expect(provenanceKey('chg-42')).toBe('ge:prov:chg-42');
    expect(provenanceRecord('chg-42', payload).key).toBe('ge:prov:chg-42');
  });

  it('is pure: same input → identical key, json, xml', () => {
    const a = provenanceRecord('chg-1', payload);
    const b = provenanceRecord('chg-1', payload);
    expect(a).toEqual(b);
  });

  it('round-trips the JSON back to the payload (plus the changeId)', () => {
    const { json } = provenanceRecord('chg-7', payload);
    const parsed = JSON.parse(json) as { changeId: string } & ProvenancePayload;
    expect(parsed.changeId).toBe('chg-7');
    const { changeId: _changeId, ...rest } = parsed;
    expect(ProvenancePayloadSchema.parse(rest)).toEqual(payload);
  });

  it('carries an optional sessionId only when present', () => {
    const withSession: ProvenancePayload = { ...payload, sessionId: asSessionId('sess-9') };
    const recordWith = provenanceRecord('c', withSession);
    expect(JSON.parse(recordWith.json).sessionId).toBe('sess-9');
    expect(recordWith.xml).toContain('sessionId="sess-9"');
    // Absent on the base payload.
    expect(provenanceRecord('c', payload).xml).not.toContain('sessionId=');
  });

  it('emits a well-formed XML part with the key and every field, escaping values', () => {
    const { xml } = provenanceRecord('chg-9', payload);
    expect(xml).toContain('key="ge:prov:chg-9"');
    expect(xml).toContain('changeId="chg-9"');
    expect(xml).toContain('agentId="review-agent@v2"');
    expect(xml).toContain('identity="v.k@acme"');
    expect(xml).toContain('contentHash="sha256:abc123"');
    expect(xml).toContain('<source title="SLA Policy" uri="https://acme/sla" locator="p.4"/>');
    expect(xml).toContain('<source title="Uptime Memo"/>');
    // Single rooted element tree, balanced tags.
    expect(xml.startsWith('<geProvenance')).toBe(true);
    expect(xml.endsWith('</geProvenance>')).toBe(true);
  });

  it('escapes XML metacharacters in titles/uris so a crafted source cannot break the part', () => {
    const hostile: ProvenancePayload = {
      ...payload,
      sources: [{ title: 'A & B <x> "q\' />', uri: 'https://acme?a=1&b=2' }],
    };
    const { xml } = provenanceRecord('c', hostile);
    expect(xml).toContain('title="A &amp; B &lt;x&gt; &quot;q&apos; /&gt;"');
    expect(xml).toContain('uri="https://acme?a=1&amp;b=2"');
    // No raw, unescaped angle bracket leaked inside an attribute value.
    expect(xml).not.toContain('<x>');
  });
});
