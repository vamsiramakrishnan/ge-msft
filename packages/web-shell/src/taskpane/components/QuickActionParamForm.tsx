import { useEffect, useRef, useState } from 'react';
import { actionParameters, fillPrompt, type QuickAction } from '@ge/contracts';
import { modeCta } from './mode-labels.js';

export interface QuickActionParamFormProps {
  /** The action whose `{{name}}` slots need values, or undefined when nothing is being collected. */
  action: QuickAction | undefined;
  disabled?: boolean;
  /** Dispatch with the collected values (keyed by parameter name). */
  onSubmit: (action: QuickAction, values: Record<string, string>) => void;
  onCancel: () => void;
}

/**
 * The fill-in step for a parameterized quick action (Workstream H). A chip whose prompt carries
 * `{{name}}` slots is NOT dispatched on click — this form collects every declared value first, so a
 * literal `{{topic}}` can never reach the model. Submit is disabled until all fields are non-empty
 * (require-values-before-dispatch); on submit the panel substitutes the values into the prompt and
 * routes the SAME typed invocation a bare chip would, with scope/intent/grounding intact.
 */
export function QuickActionParamForm({
  action,
  disabled = false,
  onSubmit,
  onCancel,
}: QuickActionParamFormProps): JSX.Element | null {
  const [values, setValues] = useState<Record<string, string>>({});
  const firstInputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (action) firstInputRef.current?.focus();
  }, [action]);
  if (!action) return null;

  const params = actionParameters(action);
  const complete = params.every((p) => (values[p.name] ?? '').trim().length > 0);

  const submit = (): void => {
    if (!complete || disabled) return;
    const trimmed: Record<string, string> = {};
    for (const p of params) trimmed[p.name] = (values[p.name] ?? '').trim();
    onSubmit(action, trimmed);
    setValues({});
  };

  const cancel = (): void => {
    setValues({});
    onCancel();
  };

  return (
    <form
      className="qa-param-form"
      data-testid="quick-action-param-form"
      aria-label={`Fill in ${action.label}`}
      onKeyDown={(e) => {
        if (e.key === 'Escape') cancel();
      }}
      onSubmit={(e) => {
        e.preventDefault();
        submit();
      }}
    >
      <div className="qa-param-title">{action.label}</div>
      {params.map((p) => (
        <label key={p.name} className="qa-param-field">
          <span className="qa-param-label">{p.label}</span>
          <input
            ref={p === params[0] ? firstInputRef : undefined}
            data-testid={`qa-param-${p.name}`}
            value={values[p.name] ?? ''}
            placeholder={p.hint ?? ''}
            autoComplete="off"
            disabled={disabled}
            onChange={(e) => setValues((prev) => ({ ...prev, [p.name]: e.target.value }))}
          />
        </label>
      ))}
      <details className="qa-prompt-preview">
        <summary>Preview instructions</summary>
        <p>{fillPrompt(action.prompt, values)}</p>
        <span>
          {action.output === 'chat'
            ? 'Answer in the conversation'
            : 'Changes will be staged for review'}
        </span>
      </details>
      <div className="qa-param-actions">
        <button type="button" className="qa-param-cancel" onClick={cancel}>
          Cancel
        </button>
        <button
          type="submit"
          className="qa-param-submit"
          data-testid="quick-action-param-submit"
          disabled={!complete || disabled}
        >
          {ctaLabel(action)}
        </button>
      </div>
    </form>
  );
}

function ctaLabel(action: QuickAction): string {
  return modeCta(action);
}
