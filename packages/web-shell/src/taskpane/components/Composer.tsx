import { useEffect, useMemo, useState } from 'react';
import type { Surface, Intent, CommandScope, GroundSource } from '@ge/contracts';
import {
  commandPaletteFor,
  type CommandVerb,
  type CommandPaletteSpec,
  type ScopeOption,
} from '@ge/contracts';

/** One typed `@`-mention: the ground kind and an optional addressable handle (e.g. a person/doc id). */
export interface ComposerMention {
  kind: GroundSource;
  ref?: string;
}

/**
 * A structured composer submit (EXPERIENCE.md §3): the chosen `/verb` intent (if any), the
 * orthogonal SCOPE the chosen segmented control resolved to, the typed `@`-ground mentions, the
 * free-text instruction (the leading `/verb` stripped, mentions left inline), and the verbatim raw.
 * One `Invocation` flows to one controller dispatch — App never re-parses a raw string.
 */
export interface ComposerInvocation {
  /** The intent from a leading `/verb` resolved against the SURFACE palette, or undefined for plain text. */
  intent?: Intent;
  /** WHERE the verb acts — the orthogonal scope (from the segmented control / a chip pre-fill). */
  scope: CommandScope;
  /** The `@`-mention tokens typed in the input, as typed {@link ComposerMention}s. */
  mentions: ComposerMention[];
  /** The free-text instruction with the leading `/verb` stripped (mentions are left inline). */
  instruction: string;
  /** The raw, verbatim input — what the user actually typed. */
  raw: string;
}

export interface ComposerProps {
  busy: boolean;
  disabled?: boolean;
  /** The current surface — scopes the `/` palette verbs offered (capability closure, ADR-0006). */
  surface?: Surface;
  /** The intents the surface can actually run; narrows the `/` palette further (ADR-0006). */
  allowedIntents?: Iterable<Intent>;
  onSend: (query: string) => void;
  onCancel: () => void;
  /** Structured submit (intent + scope + mentions + instruction). When omitted, the composer falls
   *  back to the plain `onSend` routing so existing call sites keep working unchanged. */
  onInvoke?: (invocation: ComposerInvocation) => void;
  placeholder?: string;
  /**
   * Concrete, addressable options for an addressable `GroundSource` kind (`datastore`/`document`/
   * `person`/`upload`) — e.g. `{ datastore: [{ id: dataStore.resourceName, label: dataStore.displayName }] }`
   * from `PanelController`'s discovered catalog. Typing `@datastore:` opens a SECOND picker listing
   * these, instead of leaving the mention as an unresolvable bare kind (a mention of an addressable
   * kind with no `ref` is dropped by `mentionToSelection`, never reaching grounding). A kind with no
   * entry here (or an empty list) just has no refinement step — `@kind ` alone is still accepted.
   */
  mentionOptions?: Partial<Record<GroundSource, { id: string; label: string }[]>>;
}

/** A `GroundSource` kind, as a fast membership test for the tokenizer's `@kind` recognition. */
const GROUND_KINDS = new Set<string>([
  'this',
  'unit',
  'document',
  'person',
  'datastore',
  'upload',
] satisfies GroundSource[]);

/**
 * THE one tokenizer — used for both the live affordance and submit (EXPERIENCE.md §3). Splits a raw
 * input into its leading `/verb` token (if any) and the `@`-mention tokens, resolving the verb
 * against the SURFACE palette (`spec.verbs`), NOT a cross-surface union — an out-of-scope `/verb`
 * stays `undefined` (plain text). `scope` is supplied by the segmented control, not parsed from text.
 */
export function parseComposerInput(
  raw: string,
  scope: CommandScope,
  spec?: CommandPaletteSpec,
): ComposerInvocation {
  const trimmed = raw.trim();
  let instruction = trimmed;
  let verb: string | undefined;
  const verbMatch = /^\/(\S+)\s*/.exec(trimmed);
  if (verbMatch) {
    verb = (verbMatch[1] ?? '').toLowerCase();
    instruction = trimmed.slice(verbMatch[0].length);
  }
  // `@kind` alone, or `@kind:ref` for an addressable kind picked from the refinement list (the `ref`
  // is the id `mentionToSelection` needs to resolve — see `mentionOptions` on `ComposerProps`).
  const mentions: ComposerMention[] = [];
  for (const m of trimmed.matchAll(/@(\w+)(?::(\S+))?/g)) {
    const kind = m[1]?.toLowerCase();
    if (kind === undefined || !GROUND_KINDS.has(kind)) continue;
    mentions.push({ kind: kind as GroundSource, ...(m[2] ? { ref: m[2] } : {}) });
  }
  return {
    intent: verbToIntent(verb, spec),
    scope: { ...scope },
    mentions,
    instruction,
    raw: trimmed,
  };
}

/**
 * Map a `/verb` label to its intent USING THE SURFACE PALETTE only (`spec?.verbs`). An out-of-scope
 * `/verb` (a verb another surface offers but this one doesn't) resolves to `undefined` — it is plain
 * text, not a smuggled cross-surface intent. With no spec there is no resolution (plain text).
 */
function verbToIntent(verb: string | undefined, spec?: CommandPaletteSpec): Intent | undefined {
  if (!verb || !spec) return undefined;
  const hit = spec.verbs.find((v) => v.label.replace(/^\//, '').toLowerCase() === verb);
  return hit?.intent;
}

/**
 * The union of every surface's palette verbs — kept ONLY as a "did you mean…" hint for an
 * out-of-scope `/verb` (e.g. typing `/draft` in Word). It is never used to resolve an intent; that
 * is always done against the surface palette in {@link verbToIntent}.
 */
const ALL_VERBS: CommandVerb[] = (
  ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'] as Surface[]
).flatMap((s) => commandPaletteFor(s).verbs);

/** True when `verb` is a real verb on SOME surface but not this one — the trigger for the hint. */
function offeredElsewhere(verb: string, spec?: CommandPaletteSpec): boolean {
  if (!spec) return false;
  const here = spec.verbs.some((v) => v.label.replace(/^\//, '').toLowerCase() === verb);
  if (here) return false;
  return ALL_VERBS.some((v) => v.label.replace(/^\//, '').toLowerCase() === verb);
}

/**
 * The ask box. Keyboard-submit (Enter) input with a send button that flips to Cancel while a turn
 * streams (wiring `cancel()`). A **scope segmented control** (fed by `palette.scopeOptions`, default
 * per surface) sits next to Send — scope is the orthogonal WHERE axis, never a verb. Typing a leading
 * `/` opens the surface's command palette (the `/verb` list, scoped by capability closure, ADR-0006);
 * typing `@` opens the ground-kind picker. ONE tokenizer ({@link parseComposerInput}) drives both the
 * live affordance and the submit. On submit the input becomes a typed {@link ComposerInvocation}
 * (intent + scope + mentions + instruction) handed to `onInvoke`; with no `onInvoke` it falls back to
 * plain `onSend`. Empty input is a no-op. Mode is inferred downstream — there is no agentic checkbox.
 */
export function Composer({
  busy,
  disabled = false,
  surface,
  allowedIntents,
  onSend,
  onCancel,
  onInvoke,
  placeholder,
  mentionOptions,
}: ComposerProps): JSX.Element {
  const [value, setValue] = useState('');

  const palette = useMemo(
    () => (surface ? commandPaletteFor(surface, allowedIntents) : undefined),
    [surface, allowedIntents],
  );

  const scopeOptions: ScopeOption[] = palette?.scopeOptions ?? [];
  // The selected scope index — defaults to the per-surface first option (EXPERIENCE.md §1, Tier 2).
  const [scopeIdx, setScopeIdx] = useState(0);
  const scope: CommandScope = scopeOptions[scopeIdx]?.scope ?? { kind: 'selection' };

  // The trailing token decides which affordance is open: `/…` → verb palette, `@…` → ground kinds
  // (or, once a kind + `:` is typed and that kind has options, the REFINEMENT list of concrete picks).
  const trailing = /(^|\s)([/@]\S*)$/.exec(value)?.[2];
  const showVerbs = palette !== undefined && trailing?.startsWith('/') === true;
  const isMentionToken = trailing?.startsWith('@') === true;
  const mentionBody = isMentionToken ? (trailing as string).slice(1) : '';
  const colonIdx = mentionBody.indexOf(':');
  const typedKind = (colonIdx === -1 ? mentionBody : mentionBody.slice(0, colonIdx)).toLowerCase();
  const refFilter = colonIdx === -1 ? '' : mentionBody.slice(colonIdx + 1).toLowerCase();
  const refineOptions = colonIdx !== -1 ? mentionOptions?.[typedKind as GroundSource] : undefined;
  const showRefine = isMentionToken && refineOptions !== undefined;
  const showMentions = palette !== undefined && isMentionToken && !showRefine;
  const verbFilter = showVerbs ? (trailing as string).slice(1).toLowerCase() : '';
  const mentionFilter = showMentions ? mentionBody.toLowerCase() : '';

  const verbMatches = (palette?.verbs ?? []).filter((v) =>
    v.label.replace(/^\//, '').toLowerCase().startsWith(verbFilter),
  );
  const mentionMatches = (palette?.mentionKinds ?? []).filter((k) =>
    k.toLowerCase().startsWith(mentionFilter),
  );
  const refineMatches = (refineOptions ?? []).filter(
    (o) => o.label.toLowerCase().includes(refFilter) || o.id.toLowerCase().includes(refFilter),
  );

  // The open palette and its current options — only one of the three can be open at a time.
  const paletteOpen =
    (showVerbs && verbMatches.length > 0) ||
    (showMentions && mentionMatches.length > 0) ||
    (showRefine && refineMatches.length > 0);
  const optionCount = showVerbs
    ? verbMatches.length
    : showRefine
      ? refineMatches.length
      : showMentions
        ? mentionMatches.length
        : 0;

  // Roving focus for the open palette listbox. Resets to the first option whenever the open
  // palette's option set changes (open/filter), and is clamped so it never points past the list.
  const [activeIndex, setActiveIndex] = useState(0);
  useEffect(() => {
    setActiveIndex(0);
  }, [showVerbs, showMentions, showRefine, optionCount]);
  const clampedActive = optionCount > 0 ? Math.min(activeIndex, optionCount - 1) : 0;

  /** The stable id of the option at `i` in the currently-open palette (for aria-activedescendant). */
  const optionId = (i: number): string =>
    `comp-opt-${showVerbs ? 'verb' : showRefine ? 'refine' : 'mention'}-${i}`;

  /** Apply the option at `i` of the open palette — the same effect as clicking it. */
  const selectActiveOption = (i: number): void => {
    if (showVerbs) {
      const v = verbMatches[i];
      if (v) complete(v.label);
    } else if (showRefine) {
      const o = refineMatches[i];
      if (o) complete(`@${typedKind}:${o.id}`);
    } else if (showMentions) {
      const k = mentionMatches[i];
      if (k) complete(`@${k}`);
    }
  };

  // A "did you mean…" hint: the leading /verb is a real verb on another surface but not this one.
  // (`offeredElsewhere` already encodes "not on this surface AND offered on some other surface".)
  const leadingVerb = /^\/(\S+)/.exec(value.trim())?.[1]?.toLowerCase();
  const didYouMean =
    leadingVerb !== undefined && offeredElsewhere(leadingVerb, palette) ? leadingVerb : undefined;

  /** Replace the open trailing `/`/`@` token with the picked verb/mention and keep typing. */
  const complete = (token: string): void => {
    setValue((prev) => prev.replace(/([/@]\S*)$/, `${token} `));
  };

  const submit = (): void => {
    const q = value.trim();
    if (!q || disabled || busy) return;
    if (onInvoke) onInvoke(parseComposerInput(q, scope, palette));
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
      {scopeOptions.length > 1 && (
        <div
          className="comp-scope"
          role="radiogroup"
          aria-label="Scope"
          data-testid="scope-control"
        >
          {scopeOptions.map((opt, i) => (
            <button
              key={opt.label}
              type="button"
              className="scope-option"
              role="radio"
              aria-checked={i === scopeIdx}
              data-scope-kind={opt.scope.kind}
              data-selected={i === scopeIdx ? 'true' : 'false'}
              disabled={busy || disabled}
              onClick={() => setScopeIdx(i)}
            >
              {opt.label}
            </button>
          ))}
        </div>
      )}
      {showVerbs && verbMatches.length > 0 && (
        <ul
          className="palette palette-verbs"
          role="listbox"
          aria-label="Commands"
          aria-activedescendant={optionId(clampedActive)}
        >
          {verbMatches.map((v, i) => (
            <li key={v.intent} id={optionId(i)} role="option" aria-selected={i === clampedActive}>
              <button
                type="button"
                className="palette-item"
                data-intent={v.intent}
                disabled={busy || disabled}
                onMouseEnter={() => setActiveIndex(i)}
                onClick={() => complete(v.label)}
              >
                <span className="palette-label">{v.label}</span>
                <span className="palette-desc">{v.description}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
      {didYouMean && (
        <div className="palette-hint" role="note" data-testid="verb-hint">
          {`/${didYouMean} isn't available on this surface.`}
        </div>
      )}
      {showMentions && mentionMatches.length > 0 && (
        <ul
          className="palette palette-mentions"
          role="listbox"
          aria-label="Mentions"
          aria-activedescendant={optionId(clampedActive)}
        >
          {mentionMatches.map((k, i) => (
            <li key={k} id={optionId(i)} role="option" aria-selected={i === clampedActive}>
              <button
                type="button"
                className="palette-item"
                data-mention={k}
                disabled={busy || disabled}
                onMouseEnter={() => setActiveIndex(i)}
                onClick={() => complete(`@${k}`)}
              >
                <span className="palette-label">@{k}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
      {showRefine && refineMatches.length > 0 && (
        <ul
          className="palette palette-mention-refine"
          role="listbox"
          aria-label={`Pick a ${typedKind}`}
          aria-activedescendant={optionId(clampedActive)}
        >
          {refineMatches.map((o, i) => (
            <li key={o.id} id={optionId(i)} role="option" aria-selected={i === clampedActive}>
              <button
                type="button"
                className="palette-item"
                data-mention-ref={o.id}
                disabled={busy || disabled}
                onMouseEnter={() => setActiveIndex(i)}
                onClick={() => complete(`@${typedKind}:${o.id}`)}
              >
                <span className="palette-label">{o.label}</span>
              </button>
            </li>
          ))}
        </ul>
      )}
      <div className="cb">
        <label className="visually-hidden" htmlFor="ask">
          Ask Gemini
        </label>
        <textarea
          id="ask"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (paletteOpen) {
              if (e.key === 'ArrowDown') {
                e.preventDefault();
                setActiveIndex((prev) => (optionCount > 0 ? (prev + 1) % optionCount : 0));
                return;
              }
              if (e.key === 'ArrowUp') {
                e.preventDefault();
                setActiveIndex((prev) =>
                  optionCount > 0 ? (prev - 1 + optionCount) % optionCount : 0,
                );
                return;
              }
              if (e.key === 'Enter') {
                e.preventDefault();
                selectActiveOption(clampedActive);
                return;
              }
              if (e.key === 'Escape') {
                e.preventDefault();
                // Drop the open trailing `/`/`@` token to close the palette, keeping prior text.
                setValue((prev) => prev.replace(/([/@]\S*)$/, ''));
                return;
              }
            }
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              submit();
            }
          }}
          placeholder={placeholder ?? 'Ask about the selection…'}
          autoComplete="off"
          disabled={busy || disabled}
          rows={2}
        />
        {busy ? (
          <button type="button" className="snd cancel" onClick={onCancel} aria-label="Cancel">
            <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true">
              <rect x="7" y="7" width="10" height="10" rx="2" fill="currentColor" />
            </svg>
          </button>
        ) : (
          <button
            type="submit"
            className="snd"
            aria-label="Send"
            disabled={disabled || !value.trim()}
          >
            <svg
              width="15"
              height="15"
              viewBox="0 0 24 24"
              aria-hidden="true"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="M22 2 11 13" />
              <path d="M22 2 15 22 11 13 2 9 22 2Z" />
            </svg>
          </button>
        )}
      </div>
    </form>
  );
}
