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
  if (steps.length === 0) return null;
  return (
    <section className="run-steps" aria-label="Command loop steps" aria-live="polite">
      <div className="run-steps-h eyebrow">Activity</div>
      <ol className="run-steps-list" style={{ listStyle: 'none', margin: 0, padding: 0 }}>
        {steps.map((s) => (
          <li key={s.id} className={`run-step step-${s.kind}`}>
            <span className="step-kind">{s.kind.replace(/-/g, ' ')}</span>
            <span className="step-text">{s.text}</span>
          </li>
        ))}
      </ol>
    </section>
  );
}
