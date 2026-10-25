import { useState } from 'react';
import type { PendingPlanClarification } from '../../controller.js';
import { uiStrings, type UiLanguage } from '../../ui-language.js';

export interface PlanClarificationCardProps {
  /** The conversation's language, for the card's own text. */
  language?: UiLanguage;
  pending?: PendingPlanClarification;
  disabled?: boolean;
  onAnswer: (answer: string) => void;
}

export function clarificationChoices(question: string): string[] {
  const q = question.toLowerCase();
  if (/\b(chart|visual|graph|show)\b/.test(q)) {
    return ['Hours per activity (whole week)', 'Hours by day', 'Meetings vs focus time'];
  }
  if (/\b(schedule|routine|calendar|week)\b/.test(q)) {
    return ['Google SWE schedule', 'Standard 9-5 work week', 'Fitness-heavy routine'];
  }
  if (/\b(section|scope|range|selection|where)\b/.test(q)) {
    return ['Current selection', 'Whole document', 'Choose a specific range'];
  }
  if (/\b(tone|style|format)\b/.test(q)) {
    return ['Concise', 'Executive', 'Detailed'];
  }
  return [];
}

export function PlanClarificationCard({
  pending,
  disabled = false,
  onAnswer,
  language = 'en',
}: PlanClarificationCardProps): JSX.Element | null {
  const [custom, setCustom] = useState('');
  if (!pending) return null;

  const t = uiStrings(language);
  // A correction is free text about the user's own plan: no suggested answers.
  const choices = pending.correction ? [] : unique(pending.questions.flatMap(clarificationChoices));
  const submit = (answer: string): void => {
    const trimmed = answer.trim();
    if (trimmed) onAnswer(trimmed);
  };

  return (
    <section className="command-plan clarification-card" role="group" aria-label="Clarification">
      <div className="command-plan-head">
        <span className="command-plan-title">
          {pending.correction ? t.planChangeTitle : t.clarifyTitle}
        </span>
      </div>
      {pending.previousSteps?.length ? (
        <ol className="command-plan-steps clarification-previous">
          {pending.previousSteps.map((step, idx) => (
            <li key={idx}>{step}</li>
          ))}
        </ol>
      ) : null}
      {pending.correction ? (
        <p className="clarification-ask">{pending.questions.join(' ')}</p>
      ) : (
        <ol className="command-plan-steps clarification-questions">
          {pending.questions.map((question, idx) => (
            <li key={idx}>{question}</li>
          ))}
        </ol>
      )}
      {choices.length > 0 && (
        <div className="clarification-choices" aria-label="Suggested answers">
          {choices.map((choice) => (
            <button
              key={choice}
              type="button"
              className="clarification-choice"
              disabled={disabled}
              onClick={() => submit(choice)}
            >
              {choice}
            </button>
          ))}
        </div>
      )}
      <form
        className="clarification-custom"
        onSubmit={(event) => {
          event.preventDefault();
          submit(custom);
        }}
      >
        <input
          className="field-control"
          value={custom}
          disabled={disabled}
          aria-label={t.clarifyAnswer}
          placeholder={t.clarifyAnswer}
          onChange={(event) => setCustom(event.currentTarget.value)}
        />
        <button type="submit" className="btn pr" disabled={disabled || custom.trim().length === 0}>
          {t.clarifyContinue}
        </button>
      </form>
    </section>
  );
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}
