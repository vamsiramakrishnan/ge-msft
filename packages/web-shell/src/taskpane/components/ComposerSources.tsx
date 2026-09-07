import { useEffect, useId, useRef, useState } from 'react';
import type { GroundSource } from '@ge/contracts';
import type { ComposerMention } from './Composer.js';

export type MentionOptions = Partial<Record<GroundSource, { id: string; label: string }[]>>;
export function mentionKey(mention: ComposerMention): string {
  return JSON.stringify([mention.kind, mention.ref ?? '']);
}
export function mentionLabel(mention: ComposerMention, options?: MentionOptions): string {
  if (mention.kind === 'this') return 'Current context';
  if (mention.kind === 'unit') return 'Attached sources';
  return (
    options?.[mention.kind]?.find((option) => option.id === mention.ref)?.label ??
    mention.ref ??
    mention.kind
  );
}

/** Structured picks keep resource ids out of the editable prose without losing grounding. */
export function ComposerSources({
  selected,
  options,
  disabled,
  onChange,
}: {
  selected: ComposerMention[];
  options?: MentionOptions;
  disabled: boolean;
  onChange: (mentions: ComposerMention[]) => void;
}): JSX.Element {
  const [query, setQuery] = useState('');
  const detailsRef = useRef<HTMLDetailsElement>(null);
  const close = (): void => {
    if (detailsRef.current) detailsRef.current.open = false;
    detailsRef.current?.querySelector('summary')?.focus();
  };
  useEffect(() => {
    if (disabled && detailsRef.current) detailsRef.current.open = false;
  }, [disabled]);
  const id = useId();
  const choices: ComposerMention[] = [
    ...Object.entries(options ?? {}).flatMap(([kind, items]) =>
      (items ?? []).map((item) => ({ kind: kind as GroundSource, ref: item.id })),
    ),
  ];
  const filtered = choices.filter((choice) =>
    mentionLabel(choice, options).toLowerCase().includes(query.toLowerCase()),
  );
  return (
    <details
      className="composer-sources"
      ref={detailsRef}
      onKeyDown={(event) => {
        if (event.key === 'Escape') {
          event.preventDefault();
          close();
        }
      }}
    >
      <summary aria-label="Add sources to this request">
        + Sources{selected.length > 0 ? ` (${selected.length})` : ''}
      </summary>
      <div className="composer-source-picker" role="group" aria-label="Sources for this request">
        <label htmlFor={id}>Find a source</label>
        <input
          id={id}
          type="search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder="Search connected sources"
          disabled={disabled}
        />
        <div className="composer-source-results">
          {filtered.map((choice) => {
            const checked = selected.some((item) => mentionKey(item) === mentionKey(choice));
            return (
              <button
                type="button"
                key={mentionKey(choice)}
                aria-pressed={checked}
                disabled={disabled}
                onClick={() =>
                  onChange(
                    checked
                      ? selected.filter((item) => mentionKey(item) !== mentionKey(choice))
                      : [...selected, choice],
                  )
                }
              >
                <span aria-hidden="true">{checked ? '✓' : '+'}</span>
                <span>{mentionLabel(choice, options)}</span>
              </button>
            );
          })}
          {filtered.length === 0 && (
            <p>No connected sources match. Use Context to attach host content.</p>
          )}
        </div>
        <button type="button" className="text-control" onClick={close}>
          Done selecting sources
        </button>
        <p>Connected sources apply to this request. Attach document context above.</p>
      </div>
    </details>
  );
}
