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
    <ol className="run-steps" aria-label="Command loop steps" aria-live="polite">
      {steps.map((s) => (
        <li key={s.id} className={`run-step step-${s.kind}`}>
          <span className="step-kind">{s.kind}</span>
          <span className="step-text">{s.text}</span>
        </li>
      ))}
    </ol>
  );
}
