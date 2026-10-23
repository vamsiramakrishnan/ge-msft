import { brand } from '../../brand.js';

/** How answers are shaped: kept with the pane settings, applied to every typed request. */
export interface AnswerOptionsValue {
  format?: string;
  style?: string;
}

export function AnswerOptions({
  value,
  disabled,
  onChange,
}: {
  value: AnswerOptionsValue;
  disabled?: boolean;
  onChange: (next: AnswerOptionsValue) => void;
}): JSX.Element {
  const set = (key: keyof AnswerOptionsValue, next: string): void => {
    const { [key]: _drop, ...rest } = value;
    onChange(next ? { ...rest, [key]: next } : rest);
  };
  return (
    <section className="answer-options" aria-label="Answer format and style">
      <label className="field">
        <span className="field-label">Format</span>
        <select
          className="field-control"
          aria-label="Response format"
          disabled={disabled}
          value={value.format ?? ''}
          onChange={(event) => set('format', event.target.value)}
        >
          <option value="">Let {brand.assistantName} choose</option>
          <option>Concise bullets</option>
          <option>Comparison table</option>
          <option>Decision brief</option>
          <option>Checklist with owners</option>
        </select>
      </label>
      <label className="field">
        <span className="field-label">Style</span>
        <select
          className="field-control"
          aria-label="Writing style"
          disabled={disabled}
          value={value.style ?? ''}
          onChange={(event) => set('style', event.target.value)}
        >
          <option value="">Default</option>
          <option>Plain language</option>
          <option>Executive</option>
          <option>Technical</option>
        </select>
      </label>
    </section>
  );
}
