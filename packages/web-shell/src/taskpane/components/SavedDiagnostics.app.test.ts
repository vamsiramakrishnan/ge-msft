// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { App } from './App.js';
import { makeDemoController } from '../preview-interactive.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'ge.saved-diagnostics.v1:tester';
let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
  localStorage.clear();
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
});

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

const settle = async (): Promise<void> => {
  for (let i = 0; i < 20; i++) await act(async () => Promise.resolve());
};

function typeAndSend(text: string): void {
  const box = container.querySelector<HTMLTextAreaElement>('textarea#ask')!;
  const setValue = Object.getOwnPropertyDescriptor(
    window.HTMLTextAreaElement.prototype,
    'value',
  )!.set!;
  act(() => {
    setValue.call(box, text);
    box.dispatchEvent(new Event('input', { bubbles: true }));
  });
  act(() => {
    container
      .querySelector('form.comp')!
      .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
  });
}

const saved = (): Array<{ question: string; diagnostic: Record<string, unknown> }> =>
  JSON.parse(localStorage.getItem(KEY) ?? '[]');

describe('saved diagnostics in the pane', () => {
  it('saves the typed question with its diagnostics, and never an inserted answer', async () => {
    const controller = makeDemoController('word', { pace: 0 });
    act(() => {
      root.render(createElement(App, { controller, surface: 'word', diagnosticsScope: 'tester' }));
    });

    typeAndSend('what is the launch decision?');
    await settle();
    expect(saved().map((e) => e.question)).toEqual(['what is the launch decision?']);
    expect(Object.keys(saved()[0]!.diagnostic)).toEqual(
      expect.arrayContaining(['routes', 'steps', 'errors', 'runs']),
    );

    // Clicking Insert on an answer runs its text as a program; that text must never be stored.
    let insert: Promise<void> = Promise.resolve();
    act(() => {
      insert = controller.runDirectCommands('/insert-text text="SECRET ANSWER FROM THE MODEL"');
    });
    await settle();
    act(() => controller.rejectPlan());
    await act(async () => insert);
    await settle();
    expect(controller.getState().busy).toBe(false);
    expect(localStorage.getItem(KEY)).not.toContain('SECRET ANSWER');
  });

  it('saves nothing without a signed-in scope', async () => {
    const controller = makeDemoController('word', { pace: 0 });
    act(() => {
      root.render(createElement(App, { controller, surface: 'word' }));
    });
    typeAndSend('hello');
    await settle();
    expect(localStorage.length).toBe(0);
  });
});

describe('the ⋯ earlier-diagnostics button', () => {
  it('is hidden with nothing saved, appears after a request, and copies the chosen one', async () => {
    const writeText = vi.fn(async (_text: string) => undefined);
    Object.assign(navigator, { clipboard: { writeText } });
    const controller = makeDemoController('word', { pace: 0 });
    act(() => {
      root.render(createElement(App, { controller, surface: 'word', diagnosticsScope: 'tester' }));
    });
    expect(container.querySelector('[aria-label="Earlier diagnostics"]')).toBeNull();

    typeAndSend('what is the launch decision?');
    await settle();

    const more = container.querySelector<HTMLButtonElement>('[aria-label="Earlier diagnostics"]')!;
    expect(more).not.toBeNull();
    act(() => more.click());
    const copy = container.querySelector<HTMLButtonElement>(
      'button[aria-label="Copy diagnostics for what is the launch decision?"]',
    )!;
    await act(async () => copy.click());
    expect(JSON.parse(writeText.mock.calls[0]![0])).toMatchObject({
      question: 'what is the launch decision?',
    });

    act(() => container.querySelector<HTMLButtonElement>('.saved-diagnostics-clear')!.click());
    expect(container.querySelector('[aria-label="Earlier diagnostics"]')).toBeNull();
    expect(localStorage.getItem(KEY)).toBeNull();
  });
});

describe('Activity Copy diagnostics', () => {
  it('includes the question the user typed', async () => {
    const writeText = vi.fn(async (_text: string) => undefined);
    Object.assign(navigator, { clipboard: { writeText } });
    const controller = makeDemoController('word', { pace: 0 });
    act(() => {
      root.render(createElement(App, { controller, surface: 'word', diagnosticsScope: 'tester' }));
    });

    typeAndSend('Add a total column for Q1 to Q4');
    await settle();
    act(() =>
      container.querySelector<HTMLButtonElement>('[data-testid="command-plan-confirm"]')!.click(),
    );
    await settle();
    act(() => controller.approvePlan());
    await settle();
    act(() => container.querySelector<HTMLButtonElement>('.run-steps-toggle')!.click());
    const copy = [
      ...container.querySelectorAll<HTMLButtonElement>('.run-steps-diagnostics button'),
    ].find((b) => b.textContent === 'Copy diagnostics')!;
    await act(async () => copy.click());

    const copied = JSON.parse(writeText.mock.calls[0]![0]) as {
      question: string;
      diagnostic: Record<string, unknown>;
    };
    expect(copied.question).toBe('Add a total column for Q1 to Q4');
    expect(Object.keys(copied.diagnostic)).toEqual(
      expect.arrayContaining(['routes', 'steps', 'errors', 'runs']),
    );
  });
});
