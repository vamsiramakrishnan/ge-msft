// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Composer, parseComposerInput, type ComposerProps } from './Composer.js';
import { commandPaletteFor, type CommandScope } from '@ge/contracts';

/**
 * Behavioral tests for the ask box. The real semantics: Enter submits, empty/whitespace input is a
 * no-op, the value is cleared after a submit, a submit hands a typed {@link ComposerInvocation}
 * (intent + scope + mentions + instruction) to `onInvoke` (or falls back to plain `onSend`), the
 * scope segmented control resolves the orthogonal scope, and while busy the send button becomes
 * Cancel wiring `onCancel`. There is NO agentic checkbox — mode is inferred downstream.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const SELECTION: CommandScope = { kind: 'selection' };

let container: HTMLDivElement;
let root: Root;

function render(props: Partial<ComposerProps> = {}): {
  onSend: ReturnType<typeof vi.fn>;
  onCancel: ReturnType<typeof vi.fn>;
  onInvoke: ReturnType<typeof vi.fn>;
} {
  const onSend = vi.fn();
  const onCancel = vi.fn();
  const onInvoke = vi.fn();
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(Composer, {
        busy: false,
        onSend,
        onCancel,
        onInvoke: 'onInvoke' in props ? props.onInvoke : undefined,
        ...props,
      }),
    );
  });
  return { onSend, onCancel, onInvoke };
}

function input(): HTMLTextAreaElement {
  return container.querySelector<HTMLTextAreaElement>('textarea#ask')!;
}

function type(text: string): void {
  const el = input();
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(
      window.HTMLTextAreaElement.prototype,
      'value',
    )!.set!;
    setter.call(el, text);
    el.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

function submitForm(): void {
  const form = container.querySelector('form.comp')!;
  act(() => {
    form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('Composer', () => {
  it('routes a non-empty submit to onSend when no onInvoke is given', () => {
    const { onSend } = render();
    type('what is the SLA?');
    submitForm();
    expect(onSend).toHaveBeenCalledWith('what is the SLA?');
  });

  it('trims surrounding whitespace before dispatching', () => {
    const { onSend } = render();
    type('   padded query   ');
    submitForm();
    expect(onSend).toHaveBeenCalledWith('padded query');
  });

  it('is a no-op on empty input', () => {
    const { onSend } = render();
    submitForm();
    expect(onSend).not.toHaveBeenCalled();
  });

  it('is a no-op on whitespace-only input', () => {
    const { onSend } = render();
    type('     ');
    submitForm();
    expect(onSend).not.toHaveBeenCalled();
  });

  it('clears the input after a successful submit', () => {
    render();
    type('clear me');
    submitForm();
    expect(input().value).toBe('');
  });

  it('has no agentic mode toggle (mode is inferred downstream)', () => {
    render();
    expect(container.querySelector('.mode-toggle')).toBeNull();
  });

  it('honors a custom placeholder', () => {
    render({ placeholder: 'Ask about the deck…' });
    expect(input().getAttribute('placeholder')).toBe('Ask about the deck…');
  });

  it('shows a disabled send button until there is trimmable input', () => {
    render();
    const send = () => container.querySelector<HTMLButtonElement>('button.snd[type="submit"]')!;
    expect(send().disabled).toBe(true);
    type('x');
    expect(send().disabled).toBe(false);
  });

  it('flips the send button to a Cancel control while busy and wires onCancel', () => {
    const { onCancel } = render({ busy: true });
    const cancel = container.querySelector<HTMLButtonElement>('.snd.cancel');
    expect(cancel).not.toBeNull();
    expect(container.querySelector('button[type="submit"]')).toBeNull();
    expect(input().disabled).toBe(true);
    act(() => cancel?.click());
    expect(onCancel).toHaveBeenCalledTimes(1);
  });

  it('does not submit a new turn while busy', () => {
    const { onSend } = render({ busy: true });
    input().disabled = false;
    type('queue this');
    submitForm();
    expect(onSend).not.toHaveBeenCalled();
  });
});

describe('parseComposerInput', () => {
  const word = commandPaletteFor('word');

  it('parses a leading /verb into its intent and strips it from the instruction', () => {
    const inv = parseComposerInput('/review the SLA section against policy', SELECTION, word);
    expect(inv.intent).toBe('review');
    expect(inv.instruction).toBe('the SLA section against policy');
    expect(inv.raw).toBe('/review the SLA section against policy');
    expect(inv.scope).toEqual(SELECTION);
  });

  it('resolves the /rewrite label against the surface palette to the rewrite intent', () => {
    expect(parseComposerInput('/rewrite this clause', SELECTION, word).intent).toBe('rewrite');
  });

  it('collects @-ground mentions as typed {kind} and leaves them inline in the instruction', () => {
    const inv = parseComposerInput('@this @unit compare the figures', SELECTION, word);
    expect(inv.mentions).toEqual([{ kind: 'this' }, { kind: 'unit' }]);
    expect(inv.instruction).toContain('@this @unit compare the figures');
  });

  it('ignores @-tokens that are not a known ground kind', () => {
    const inv = parseComposerInput('@this @bob hello', SELECTION, word);
    expect(inv.mentions).toEqual([{ kind: 'this' }]);
  });

  it('captures a @kind:ref mention (picked from the catalog refinement list) with its id', () => {
    const inv = parseComposerInput(
      '@datastore:sp-docs summarize the risk register',
      SELECTION,
      word,
    );
    expect(inv.mentions).toEqual([{ kind: 'datastore', ref: 'sp-docs' }]);
  });

  it('returns no intent for a plain question and keeps the full instruction', () => {
    const inv = parseComposerInput('what is the renewal date?', SELECTION, word);
    expect(inv.intent).toBeUndefined();
    expect(inv.instruction).toBe('what is the renewal date?');
    expect(inv.mentions).toEqual([]);
  });

  it('leaves an unknown /verb intentless rather than inventing one', () => {
    expect(parseComposerInput('/bogus do a thing', SELECTION, word).intent).toBeUndefined();
  });

  it('does NOT resolve an out-of-scope /verb against a cross-surface union', () => {
    // /draft is a real verb on PowerPoint/OneNote/Outlook but NOT on Word — it must stay plain text.
    expect(parseComposerInput('/draft a section', SELECTION, word).intent).toBeUndefined();
    expect(
      parseComposerInput('/draft a section', SELECTION, commandPaletteFor('powerpoint')).intent,
    ).toBe('draft');
  });

  it('carries the supplied scope onto the invocation verbatim', () => {
    const doc: CommandScope = { kind: 'document' };
    expect(parseComposerInput('/review it', doc, word).scope).toEqual(doc);
  });
});

describe('Composer / and @ palette + scope control + structured submit', () => {
  it('offers the surfaces /verb palette when the input opens with a slash', () => {
    render({ surface: 'word' });
    type('/');
    const items = [...container.querySelectorAll('.palette-verbs .palette-label')].map(
      (e) => e.textContent,
    );
    expect(items).toEqual(commandPaletteFor('word').verbs.map((v) => v.label));
  });

  it('filters the /verb palette by the typed prefix', () => {
    render({ surface: 'word' });
    type('/rev');
    const items = [...container.querySelectorAll('.palette-verbs .palette-label')].map(
      (e) => e.textContent,
    );
    expect(items).toEqual(['/review']);
  });

  it('shows a did-you-mean hint for a verb offered only on another surface', () => {
    render({ surface: 'word' });
    type('/draft ');
    const hint = container.querySelector('[data-testid="verb-hint"]');
    expect(hint?.textContent).toContain('/draft');
  });

  it('offers the @-ground kinds when the trailing token starts with @', () => {
    render({ surface: 'word' });
    type('summarize @');
    const kinds = [...container.querySelectorAll('.palette-mentions .palette-label')].map(
      (e) => e.textContent,
    );
    expect(kinds).toEqual(['@this', '@unit', '@datastore']);
  });

  it('opens the catalog refinement list once a configured kind + colon is typed', () => {
    render({
      surface: 'word',
      mentionOptions: {
        datastore: [
          { id: 'sp-docs', label: 'SharePoint Docs' },
          { id: 'sp-hr', label: 'HR Policies' },
        ],
      },
    });
    type('@datastore:');
    const labels = [...container.querySelectorAll('.palette-mention-refine .palette-label')].map(
      (e) => e.textContent,
    );
    expect(labels).toEqual(['SharePoint Docs', 'HR Policies']);
    // The bare-kind picker is closed while refining — only one palette shows at a time.
    expect(container.querySelector('.palette-mentions')).toBeNull();
  });

  it('filters the refinement list by what is typed after the colon', () => {
    render({
      surface: 'word',
      mentionOptions: {
        datastore: [
          { id: 'sp-docs', label: 'SharePoint Docs' },
          { id: 'sp-hr', label: 'HR Policies' },
        ],
      },
    });
    type('@datastore:hr');
    const labels = [...container.querySelectorAll('.palette-mention-refine .palette-label')].map(
      (e) => e.textContent,
    );
    expect(labels).toEqual(['HR Policies']);
  });

  it('picking a refinement option writes the addressable @kind:id token into the input', () => {
    render({
      surface: 'word',
      mentionOptions: { datastore: [{ id: 'sp-docs', label: 'SharePoint Docs' }] },
    });
    type('@datastore:');
    const item = container.querySelector<HTMLButtonElement>(
      '.palette-mention-refine .palette-item',
    )!;
    act(() => item.click());
    expect(input().value).toBe('@datastore:sp-docs ');
  });

  it('does not open a refinement list for a kind with no configured options', () => {
    render({ surface: 'word', mentionOptions: { datastore: [] } });
    type('@datastore:');
    expect(container.querySelector('.palette-mention-refine')).toBeNull();
  });

  it('completing a verb writes the /verb token into the input', () => {
    render({ surface: 'word' });
    type('/rev');
    const review = [
      ...container.querySelectorAll<HTMLButtonElement>('.palette-verbs .palette-item'),
    ].find((b) => b.textContent?.includes('/review'))!;
    act(() => review.click());
    expect(input().value).toBe('/review ');
  });

  it('renders the surface scope segmented control with the default selected first', () => {
    render({ surface: 'word' });
    const control = container.querySelector('[data-testid="scope-control"]')!;
    const opts = [...control.querySelectorAll<HTMLButtonElement>('.scope-option')];
    const expected = commandPaletteFor('word').scopeOptions;
    expect(opts.map((o) => o.textContent)).toEqual(expected.map((o) => o.label));
    expect(opts[0]?.getAttribute('data-selected')).toBe('true');
  });

  it('picking a scope sets it on the submitted invocation', () => {
    const onInvoke = vi.fn();
    render({ surface: 'word', onInvoke });
    const opts = [
      ...container.querySelectorAll<HTMLButtonElement>(
        '[data-testid="scope-control"] .scope-option',
      ),
    ];
    const docIdx = commandPaletteFor('word').scopeOptions.findIndex(
      (o) => o.scope.kind === 'document',
    );
    act(() => opts[docIdx]?.click());
    type('/review the doc');
    submitForm();
    expect(onInvoke.mock.calls[0]![0].scope.kind).toBe('document');
  });

  it('routes a structured submit through onInvoke with the parsed intent + typed mentions', () => {
    const onInvoke = vi.fn();
    const { onSend } = render({ surface: 'word', onInvoke });
    type('/review @this the clause');
    submitForm();
    expect(onInvoke).toHaveBeenCalledTimes(1);
    const inv = onInvoke.mock.calls[0]![0];
    expect(inv.intent).toBe('review');
    expect(inv.mentions).toEqual([{ kind: 'this' }]);
    expect(inv.raw).toBe('/review @this the clause');
    expect(onSend).not.toHaveBeenCalled();
  });
});
