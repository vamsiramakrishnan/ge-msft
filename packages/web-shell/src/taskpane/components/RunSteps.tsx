import { useId, useState } from 'react';
import type { RunStep } from '../../controller.js';

export interface RunStepsProps {
  steps: RunStep[];
}

/**
 * The command-loop transcript: a compact, ordered list of the `CommandLoopEvent`s
 * (turn / command / read / write / done) so the user can watch the read-many/write-one loop
 * progress. Rendered as an ordered list with a polite live region so new steps are announced.
 */
export function RunSteps({ steps }: RunStepsProps): JSX.Element | null {
  const [expanded, setExpanded] = useState(false);
  const listId = useId();

  if (steps.length === 0) return null;

  const latest = steps[steps.length - 1];
  const stepCount = steps.length === 1 ? '1 step' : `${steps.length} steps`;

  return (
    <section
      className={`run-steps${expanded ? ' is-expanded' : ''}`}
      aria-label="Command loop steps"
      aria-live="polite"
    >
      <button
        type="button"
        className="run-steps-toggle"
        aria-expanded={expanded}
        aria-controls={listId}
        onClick={() => setExpanded((value) => !value)}
      >
        <span className="run-steps-title eyebrow">Activity</span>
        <span className="run-steps-summary">
          <span className="run-steps-count">{stepCount}</span>
          {latest ? (
            <span className="run-steps-latest">
              <span className="step-kind">{latest.kind.replace(/-/g, ' ')}</span>
              <span className="step-text">{latest.text}</span>
            </span>
          ) : null}
        </span>
        <span className="run-steps-caret" aria-hidden="true" />
      </button>
      <ol
        id={listId}
        className="run-steps-list"
        hidden={!expanded}
        style={{ listStyle: 'none', margin: 0, padding: 0 }}
      >
        {steps.map((s) => (
          <li key={s.id} className={`run-step step-${s.kind}`}>
            <span className="step-kind">{s.kind.replace(/-/g, ' ')}</span>
            <span className="step-body">
              <span className="step-text">{s.text}</span>
              {s.artifact ? <WorkspaceArtifactCard artifact={s.artifact} /> : null}
            </span>
          </li>
        ))}
      </ol>
    </section>
  );
}

function WorkspaceArtifactCard({
  artifact,
}: {
  artifact: NonNullable<RunStep['artifact']>;
}): JSX.Element {
  return (
    <article className="workspace-artifact-card">
      <header className="workspace-artifact-head">
        <span className="workspace-artifact-title">{artifact.title}</span>
      </header>
      {artifact.meta.length > 0 ? (
        <ul className="workspace-artifact-meta" aria-label="Artifact metadata">
          {artifact.meta.map((item) => (
            <li key={item}>{item}</li>
          ))}
        </ul>
      ) : null}
      {artifact.preview ? (
        <pre className="workspace-artifact-preview">{artifact.preview}</pre>
      ) : null}
      {artifact.matches && artifact.matches.length > 0 ? (
        <ol className="workspace-artifact-matches" aria-label="Artifact matches">
          {artifact.matches.map((match) => (
            <li key={`${match.line}:${match.text}`}>
              <span className="workspace-artifact-line">L{match.line}</span>
              <span>{match.text}</span>
            </li>
          ))}
        </ol>
      ) : null}
    </article>
  );
}
