// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { Composer, type ComposerProps } from './Composer.js';

/**
 * Behavioral tests for the ask box. The real semantics: Enter submits, empty/whitespace input is a
 * no-op, the value is cleared after a submit, the mode toggle routes a submit to `onRun` (agentic
 * read-many/write-one loop) vs `onSend` (grounded chat), and while busy the send button becomes
 * Cancel wiring `onCancel`. The busy flag also disables the mode toggle.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(props: Partial<ComposerProps> = {}): {
  onSend: ReturnType<typeof vi.fn>;
  onRun: ReturnType<typeof vi.fn>;
  onCancel: ReturnType<typeof vi.fn>;
} {
  const onSend = vi.fn();
  const onRun = vi.fn();
  const onCancel = vi.fn();
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(Composer, {
        busy: false,
        onSend,
        onRun,
        onCancel,
        ...props,
      }),
    );
  });
  return { onSend, onRun, onCancel };
}

function input(): HTMLInputElement {
  return container.querySelector<HTMLInputElement>('input#ask')!;
}

function type(text: string): void {
  const el = input();
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype,
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

function toggleAgentic(): void {
  const cb = container.querySelector<HTMLInputElement>('.mode-toggle input')!;
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype,
      'checked',
    )!.set!;
    setter.call(cb, !cb.checked);
    cb.dispatchEvent(new Event('click', { bubbles: true }));
    cb.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('Composer', () => {
  it('routes a non-empty submit to onSend in grounded (default) mode', () => {
    const { onSend, onRun } = render();
    type('what is the SLA?');
    submitForm();
    expect(onSend).toHaveBeenCalledWith('what is the SLA?');
    expect(onRun).not.toHaveBeenCalled();
  });

  it('trims surrounding whitespace before dispatching', () => {
    const { onSend } = render();
    type('   padded query   ');
    submitForm();
    expect(onSend).toHaveBeenCalledWith('padded query');
  });

  it('is a no-op on empty input', () => {
    const { onSend, onRun } = render();
    submitForm();
    expect(onSend).not.toHaveBeenCalled();
    expect(onRun).not.toHaveBeenCalled();
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

  it('routes the submit to onRun when agentic mode is enabled', () => {
    const { onSend, onRun } = render();
    toggleAgentic();
    type('build the exposure model');
    submitForm();
    expect(onRun).toHaveBeenCalledWith('build the exposure model');
    expect(onSend).not.toHaveBeenCalled();
  });

  it('swaps the agentic placeholder and label when the mode toggle is on', () => {
    render();
    expect(input().getAttribute('placeholder')).toContain('Ask about the selection');
    toggleAgentic();
    expect(input().getAttribute('placeholder')).toContain('it will read, then propose writes');
    expect(container.querySelector('label[for="ask"]')?.textContent).toContain(
      'Give Gemini a task',
    );
  });

  it('honors a custom placeholder in grounded mode', () => {
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
    act(() => cancel?.click());
    expect(onCancel).toHaveBeenCalledTimes(1);
  });

  it('disables the mode toggle while a turn is busy', () => {
    render({ busy: true });
    expect(container.querySelector<HTMLInputElement>('.mode-toggle input')?.disabled).toBe(true);
  });
});
