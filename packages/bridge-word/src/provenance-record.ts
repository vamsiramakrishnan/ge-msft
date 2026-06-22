import type { ProvenancePayload } from '@ge/contracts';

/**
 * A serialized durable-provenance record for one write, keyed by the write's `changeId`.
 *
 * Every reversible agent write should leave a traceable record in the host's durable metadata
 * (CLAUDE.md: "reversible, provenanced writes … each carrying agent id, sources, identity,
 * timestamp, and a content hash in the host's durable metadata"). This is the *pure* serializer
 * — it turns a {@link ProvenancePayload} into a stable key plus both a JSON and an OOXML custom-XML
 * representation, so each surface can persist whichever its host supports (Word → a custom XML
 * part via `customXmlParts.add`; Excel → workbook settings as JSON). It NEVER touches Office.js,
 * so it's unit-tested. (The same tiny helper is duplicated per bridge — neither imports the other.)
 */
export interface ProvenanceRecord {
  /** Stable, idempotent settings/part key: `ge:prov:<changeId>`. Re-applying overwrites. */
  readonly key: string;
  /** A well-formed OOXML custom-XML part carrying the record (used by Word). */
  readonly xml: string;
  /** A round-trippable JSON serialization of the record (used by Excel settings). */
  readonly json: string;
}

/** The stable namespace + key for a provenance record, derived only from the `changeId`. */
export function provenanceKey(changeId: string): string {
  return `ge:prov:${changeId}`;
}

/**
 * Serialize a {@link ProvenancePayload} into a durable {@link ProvenanceRecord}. Pure: the same
 * payload + changeId always yields the same key, JSON, and XML. The JSON round-trips back to the
 * payload (plus the `changeId` we fold in); the XML is well-formed (every value escaped) so a
 * crafted source title/uri can't break the part.
 */
export function provenanceRecord(changeId: string, payload: ProvenancePayload): ProvenanceRecord {
  const key = provenanceKey(changeId);
  const record = { changeId, ...payload };
  return {
    key,
    json: JSON.stringify(record),
    xml: toXml(key, record),
  };
}

/** Render the record as a well-formed custom-XML part (a single namespaced element tree). */
function toXml(key: string, record: { changeId: string } & ProvenancePayload): string {
  const sources = record.sources
    .map(
      (s) =>
        `<source title="${esc(s.title)}"` +
        (s.uri !== undefined ? ` uri="${esc(s.uri)}"` : '') +
        (s.locator !== undefined ? ` locator="${esc(s.locator)}"` : '') +
        '/>',
    )
    .join('');
  return (
    `<geProvenance xmlns="https://gemini.google/ge/provenance" key="${esc(key)}" ` +
    `changeId="${esc(record.changeId)}" agentId="${esc(record.agentId)}" ` +
    `identity="${esc(record.identity)}" timestamp="${esc(record.timestamp)}" ` +
    `contentHash="${esc(record.contentHash)}"` +
    (record.sessionId !== undefined ? ` sessionId="${esc(record.sessionId)}"` : '') +
    `><sources>${sources}</sources></geProvenance>`
  );
}

/** Escape the five XML predefined entities so any attribute value stays well-formed. */
function esc(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
