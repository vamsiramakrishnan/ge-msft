import { useMemo, useState } from 'react';
import type { Surface, Intent } from '@ge/contracts';
import { commandPaletteFor, type CommandVerb, type CommandPaletteSpec } from '@ge/contracts';

/** A structured composer submit: the chosen `/verb` intent (if any), `@`-mentions, and free text. */
export interface ComposerInvocation {
  /** The intent from a leading `/verb`, or undefined for a plain question. */
  intent?: Intent;
  /** The `@`-mention tokens typed in the input (the bare kinds/handles, without the `@`). */
  mentions: string[];
  /** The free text with the leading `/verb` stripped (mentions are left inline). */
  text: string;
  /** The raw, verbatim input — what the user actually typed. */
  raw: string;
}

export interface ComposerProps {
  busy: boolean;
  /** The current surface — scopes the `/` palette verbs offered (capability closure, ADR-0006). */
  surface?: Surface;
  /** The intents the surface can actually run; narrows the `/` palette further (ADR-0006). */
  allowedIntents?: Iterable<Intent>;
  onSend: (query: string) => void;
  /** Start the ADR-0004 read-many/write-one command loop (agentic mode). */
  onRun: (task: string) => void;
  onCancel: () => void;
  /** Structured submit (intent + mentions + text). When omitted, the composer falls back to the
   *  plain `onSend`/`onRun` routing so existing call sites keep working unchanged. */
  onInvoke?: (invocation: ComposerInvocation) => void;
  placeholder?: string;
}

/** Split a composer input into its leading `/verb` (if any), `@`-mentions, and stripped text. */
export function parseComposerInput(raw: string): ComposerInvocation {
  const trimmed = raw.trim();
  let text = trimmed;
  let verb: string | undefined;
  const verbMatch = /^\/(\S+)\s*/.exec(trimmed);
  if (verbMatch) {
    verb = (verbMatch[1] ?? '').toLowerCase();
    text = trimmed.slice(verbMatch[0].length);
  }
  const mentions = [...trimmed.matchAll(/@(\w+)/g)]
    .map((m) => m[1])
    .filter((m): m is string => m !== undefined);
  return { intent: verbToIntent(verb), mentions, text, raw: trimmed };
}

/** Map a `/verb` label (the user types `/review`, the spec carries `/review`) to its intent. */
function verbToIntent(verb: string | undefined, spec?: CommandPaletteSpec): Intent | undefined {
  if (!verb) return undefined;
  const verbs: CommandVerb[] = spec?.verbs ?? ALL_VERBS;
  const hit = verbs.find((v) => v.label.replace(/^\//, '').toLowerCase() === verb);
  return hit?.intent;
}

/** The union of every surface's palette verbs — used for label→intent lookup when no spec is given. */
const ALL_VERBS: CommandVerb[] = (
  ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as Surface[]
).flatMap((s) => commandPaletteFor(s).verbs);

/**
 * The ask box: keyboard-submit (Enter) input with a send button that flips to Cancel while a turn
 * is streaming, wiring the controller's `cancel()`. A mode toggle switches between grounded chat
 * (`send`) and the agentic command loop (`run`, ADR-0004). Typing a leading `/` opens the
 * surface's command palette (the `/verb` list, ADR-0004 grammar, scoped by capability closure,
 * ADR-0006); typing `@` opens the mention-kind picker. On submit the input is parsed into a
 * structured {@link ComposerInvocation} (intent + mentions + free text) and handed to `onInvoke`
 * when present — otherwise it falls back to the plain `onSend`/`onRun` routing. Empty input is a
 * no-op (matches controller).
 */
export function Composer({
  busy,
  surface,
  allowedIntents,
  onSend,
  onRun,
  onCancel,
  onInvoke,
  placeholder,
}: ComposerProps): JSX.Element {
  const [value, setValue] = useState('');
  const [agentic, setAgentic] = useState(false);

  const palette = useMemo(
    () => (surface ? commandPaletteFor(surface, allowedIntents) : undefined),
    [surface, allowedIntents],
  );

  // The trailing token decides which affordance is open: `/…` → verb palette, `@…` → mention kinds.
  const trailing = /(^|\s)([/@]\S*)$/.exec(value)?.[2];
  const showVerbs = palette !== undefined && trailing?.startsWith('/') === true;
  const showMentions = palette !== undefined && trailing?.startsWith('@') === true;
  const verbFilter = showVerbs ? (trailing as string).slice(1).toLowerCase() : '';
  const mentionFilter = showMentions ? (trailing as string).slice(1).toLowerCase() : '';

  const verbMatches = (palette?.verbs ?? []).filter((v) =>
    v.label.replace(/^\//, '').toLowerCase().startsWith(verbFilter),
  );
  const mentionMatches = (palette?.mentionKinds ?? []).filter((k) =>
    k.toLowerCase().startsWith(mentionFilter),
  );

  /** Replace the open trailing `/`/`@` token with the picked verb/mention and keep typing. */
  const complete = (token: string): void => {
    setValue((prev) => prev.replace(/([/@]\S*)$/, `${token} `));
  };

  const submit = (): void => {
    const q = value.trim();
    if (!q) return;
    if (onInvoke) onInvoke(parseComposerInput(q));
    else if (agentic) onRun(q);
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
      {showVerbs && verbMatches.length > 0 && (
        <ul className="palette palette-verbs" role="listbox" aria-label="Commands">
          {verbMatches.map((v) => (
            <li key={v.intent} role="option" aria-selected="false">
              <button
                type="button"
                className="palette-item"
                data-intent={v.intent}
                onClick={() => complete(v.label)}
              >
                <span className="palette-label">{v.label}</span>
                <span className="palette-desc">{v.description}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
      {showMentions && mentionMatches.length > 0 && (
        <ul className="palette palette-mentions" role="listbox" aria-label="Mentions">
          {mentionMatches.map((k) => (
            <li key={k} role="option" aria-selected="false">
              <button
                type="button"
                className="palette-item"
                data-mention={k}
                onClick={() => complete(`@${k}`)}
              >
                <span className="palette-label">@{k}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
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
