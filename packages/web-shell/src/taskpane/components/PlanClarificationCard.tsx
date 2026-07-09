import { useState } from 'react';
import type { PendingPlanClarification } from '../../controller.js';

export interface PlanClarificationCardProps {
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
}: PlanClarificationCardProps): JSX.Element | null {
  const [custom, setCustom] = useState('');
  if (!pending) return null;

  const choices = unique(pending.questions.flatMap(clarificationChoices));
  const submit = (answer: string): void => {
    const trimmed = answer.trim();
    if (trimmed) onAnswer(trimmed);
  };

  return (
    <section className="command-plan clarification-card" role="group" aria-label="Clarification">
      <div className="command-plan-head">
        <span className="command-plan-verb">clarify</span>
        <span className="command-plan-tag">answer to continue planning</span>
      </div>
      <ol className="command-plan-steps clarification-questions">
        {pending.questions.map((question, idx) => (
          <li key={idx}>{question}</li>
        ))}
      </ol>
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
          value={custom}
          disabled={disabled}
          placeholder="Type a custom answer"
          onChange={(event) => setCustom(event.currentTarget.value)}
        />
        <button type="submit" disabled={disabled || custom.trim().length === 0}>
          Continue
        </button>
      </form>
    </section>
  );
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}
