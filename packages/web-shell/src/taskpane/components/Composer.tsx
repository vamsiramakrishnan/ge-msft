import { useState } from 'react';

export interface ComposerProps {
  busy: boolean;
  onSend: (query: string) => void;
  onCancel: () => void;
  placeholder?: string;
}

/**
 * The ask box: keyboard-submit (Enter) input with a send button that flips to Cancel while a turn
 * is streaming, wiring the controller's `cancel()`. Empty input is a no-op (matches controller).
 */
export function Composer({ busy, onSend, onCancel, placeholder }: ComposerProps): JSX.Element {
  const [value, setValue] = useState('');

  const submit = (): void => {
    const q = value.trim();
    if (!q) return;
    onSend(q);
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
      <div className="cb">
        <label className="visually-hidden" htmlFor="ask">
          Ask Gemini
        </label>
        <input
          id="ask"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder={placeholder ?? 'Ask about the selection…'}
          autoComplete="off"
        />
        {busy ? (
          <button type="button" className="snd cancel" onClick={onCancel} aria-label="Cancel">
            ◼
          </button>
        ) : (
          <button type="submit" className="snd" aria-label="Send" disabled={!value.trim()}>
            →
          </button>
        )}
      </div>
    </form>
  );
}
