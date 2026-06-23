import type { ProvenancePayload } from '@ge/contracts';

export interface ProvenanceDetailProps {
  provenance: ProvenancePayload;
}

/**
 * Only render an http(s) source link; anything else is inert text so an untrusted source can't
 * smuggle an executable href (mirrors `MessageThread.safeHttpUri`).
 */
function safeHttpUri(uri: string | undefined): string | undefined {
  if (!uri) return undefined;
  try {
    const parsed = new URL(uri, window.location.origin);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:' ? parsed.href : undefined;
  } catch {
    return undefined;
  }
}

/** A best-effort, locale-stable rendering of an ISO timestamp; falls back to the raw string. */
function formatTimestamp(iso: string): string {
  const t = Date.parse(iso);
  return Number.isNaN(t) ? iso : new Date(t).toLocaleString();
}

/**
 * The provenance drill-down for an applied write: who (agent + signed-in identity), when, the
 * grounding sources, and the content hash — the "traceable + reversible" record the host also
 * stamps into durable metadata. Presentational only; it reflects a `ProvenancePayload` the turn
 * carried and never mutates anything. Rendered as a definition list for screen-reader pairing.
 */
export function ProvenanceDetail({ provenance }: ProvenanceDetailProps): JSX.Element {
  const { agentId, identity, timestamp, sources, contentHash } = provenance;
  return (
    <dl className="provenance" aria-label="Change provenance">
      <div className="prov-row">
        <dt>Agent</dt>
        <dd className="mono">{agentId}</dd>
      </div>
      <div className="prov-row">
        <dt>Identity</dt>
        <dd>{identity}</dd>
      </div>
      <div className="prov-row">
        <dt>When</dt>
        <dd>
          <time dateTime={timestamp}>{formatTimestamp(timestamp)}</time>
        </dd>
      </div>
      <div className="prov-row">
        <dt>Sources</dt>
        <dd>
          {sources.length === 0 ? (
            <span className="muted small">No sources recorded.</span>
          ) : (
            <ul className="prov-sources">
              {sources.map((s, i) => {
                const href = safeHttpUri(s.uri);
                const label = s.locator ? `${s.title} · ${s.locator}` : s.title;
                return (
                  <li key={`${s.title}-${i}`}>
                    {href ? (
                      <a href={href} target="_blank" rel="noreferrer noopener">
                        {label}
                      </a>
                    ) : (
                      <span>{label}</span>
                    )}
                  </li>
                );
              })}
            </ul>
          )}
        </dd>
      </div>
      <div className="prov-row">
        <dt>Content hash</dt>
        <dd className="mono prov-hash">{contentHash}</dd>
      </div>
    </dl>
  );
}
