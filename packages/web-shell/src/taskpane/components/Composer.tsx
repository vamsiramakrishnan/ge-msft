import { useState } from 'react';

export interface ComposerProps {
  busy: boolean;
  onSend: (query: string) => void;
  /** Start the ADR-0004 read-many/write-one command loop (agentic mode). */
  onRun: (task: string) => void;
  onCancel: () => void;
  placeholder?: string;
}

/**
 * The ask box: keyboard-submit (Enter) input with a send button that flips to Cancel while a turn
 * is streaming, wiring the controller's `cancel()`. A mode toggle switches between grounded chat
 * (`send`) and the agentic command loop (`run`, ADR-0004). Empty input is a no-op (matches
 * controller).
 */
export function Composer({
  busy,
  onSend,
  onRun,
  onCancel,
  placeholder,
}: ComposerProps): JSX.Element {
  const [value, setValue] = useState('');
  const [agentic, setAgentic] = useState(false);

  const submit = (): void => {
    const q = value.trim();
    if (!q) return;
    if (agentic) onRun(q);
    else onSend(q);
    setValue('');
  };

  return (
    <form
      className="comp"
      onSubmit={(e) => {
        e.preventDefault();
        submit();
      }}
    >
      <div className="comp-mode">
        <label className="mode-toggle">
          <input
            type="checkbox"
            checked={agentic}
            onChange={(e) => setAgentic(e.target.checked)}
            disabled={busy}
          />
          <span>Agentic (read & propose writes)</span>
        </label>
      </div>
      <div className="cb">
        <label className="visually-hidden" htmlFor="ask">
          {agentic ? 'Give Gemini a task' : 'Ask Gemini'}
        </label>
        <input
          id="ask"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder={
            agentic
              ? 'Give a task (it will read, then propose writes)…'
              : (placeholder ?? 'Ask about the selection…')
          }
          autoComplete="off"
        />
        {busy ? (
          <button type="button" className="snd cancel" onClick={onCancel} aria-label="Cancel">
            ◼
          </button>
        ) : (
          <button
            type="submit"
            className="snd"
            aria-label={agentic ? 'Run task' : 'Send'}
            disabled={!value.trim()}
          >
            →
          </button>
        )}
      </div>
    </form>
  );
}
