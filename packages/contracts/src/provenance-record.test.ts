import { describe, expect, it } from 'vitest';
import { asSessionId } from './brand.js';
import { provenanceKey, provenanceRecord } from './provenance-record.js';

describe('shared durable provenance serialization', () => {
  it('preserves exact JSON and XML bytes for the existing persisted representation', () => {
    const record = provenanceRecord('change-1', {
      agentId: 'agent',
      identity: 'owner',
      timestamp: '2026-09-08T00:00:00Z',
      contentHash: 'sha256:abc',
      sources: [{ title: 'Policy', uri: 'https://example.com/policy', locator: 'A1' }],
      sessionId: asSessionId('session-1'),
    });
    expect(record).toEqual({
      key: 'ge:prov:change-1',
      json: '{"changeId":"change-1","agentId":"agent","identity":"owner","timestamp":"2026-09-08T00:00:00Z","contentHash":"sha256:abc","sources":[{"title":"Policy","uri":"https://example.com/policy","locator":"A1"}],"sessionId":"session-1"}',
      xml: '<geProvenance xmlns="https://gemini.google/ge/provenance" key="ge:prov:change-1" changeId="change-1" agentId="agent" identity="owner" timestamp="2026-09-08T00:00:00Z" contentHash="sha256:abc" sessionId="session-1"><sources><source title="Policy" uri="https://example.com/policy" locator="A1"/></sources></geProvenance>',
    });
  });
  it('escapes all XML attribute metacharacters and preserves optional-field omission', () => {
    const payload = {
      agentId: 'agent',
      identity: 'owner',
      timestamp: 'timestamp',
      contentHash: 'hash',
      sources: [{ title: `A & B <x> "quote" 'apostrophe'` }],
    };
    expect(provenanceKey('id')).toBe('ge:prov:id');
    const record = provenanceRecord('id', payload);
    expect(record.xml).toContain(
      'title="A &amp; B &lt;x&gt; &quot;quote&quot; &apos;apostrophe&apos;"',
    );
    expect(record.xml).not.toContain('sessionId=');
    expect(record.xml).not.toContain(' uri=');
    expect(JSON.parse(record.json)).toEqual({ changeId: 'id', ...payload });
  });
});
