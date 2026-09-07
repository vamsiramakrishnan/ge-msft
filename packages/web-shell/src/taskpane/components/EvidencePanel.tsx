import type { EvidenceState } from '@ge/runtime';
export function EvidencePanel({ evidence }: { evidence?: EvidenceState }): JSX.Element | null {
  if (!evidence || evidence.status === 'idle') return null;
  return (
    <details className="evidence-panel">
      <summary>
        <strong>Evidence</strong>
        <span>
          {evidence.sources.length} sources
          {evidence.score !== undefined
            ? ` · ${Math.round(evidence.score * 100)}% support`
            : ` · ${evidence.status}`}
        </span>
      </summary>
      <p>{evidence.message}</p>
      <ul>
        {evidence.sources.map((s) => (
          <li key={s.id}>
            {s.uri ? (
              <a href={s.uri} target="_blank" rel="noopener noreferrer">
                {s.title}
              </a>
            ) : (
              s.title
            )}
          </li>
        ))}
      </ul>
    </details>
  );
}
