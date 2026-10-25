// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { App } from './App.js';
import { makeDemoController } from '../preview-interactive.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

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

function choose(selector: string, value: string): void {
  const select = container.querySelector<HTMLSelectElement>(selector)!;
  act(() => {
    select.value = value;
    select.dispatchEvent(new Event('change', { bubbles: true }));
  });
}

describe('answer format and style in Settings', () => {
  it('lives in the sidebar Settings, not under the ask box, and shapes the next request', async () => {
    const controller = makeDemoController('word', { pace: 0 });
    act(() => {
      root.render(createElement(App, { controller, surface: 'word' }));
    });
    // No catalog client: Settings still opens, because it holds the answer options.
    expect(container.querySelector('form.comp [aria-label="Response format"]')).toBeNull();
    const settings = container.querySelector<HTMLButtonElement>('[data-panel="settings"]')!;
    expect(settings).not.toBeNull();
    act(() => settings.click());
    expect(container.querySelector('.tw-modal-title')?.textContent).toContain('Settings');

    choose('[aria-label="Response format"]', 'Decision brief');
    choose('[aria-label="Writing style"]', 'Executive');
    act(() => container.querySelector<HTMLButtonElement>('.tw-modal-close')!.click());

    typeAndSend('what is the launch decision?');
    await settle();
    expect(container.querySelector('.thread')?.textContent).toContain(
      'Output format: Decision brief. Writing style: Executive.',
    );
  });
});
