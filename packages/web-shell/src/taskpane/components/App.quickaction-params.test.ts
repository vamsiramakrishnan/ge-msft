// @vitest-environment jsdom
import { describe, it, expect, vi, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ChangeId,
  ContextRef,
  SseEvent,
} from '@ge/contracts';
import type { CommandLoopEvent, RunCommandsOptions } from '@ge/runtime';
import type { HostEvent } from '@ge/triggers';
import { PanelController, type AssistLike, type ContextLister } from '../../controller.js';
import { App } from './App.js';

/**
 * Workstream H — a parameterized quick action (a `{{name}}` slot in its prompt) must NOT dispatch on
 * click: the panel opens a fill form, collects the value, and only then routes the typed invocation
 * with the slot substituted. A literal `{{…}}` can never reach the model (fail-closed guard).
 */

class RecordingAssist implements AssistLike {
  context = { size: 0 };
  runTasks: string[] = [];
  attachRef(): Promise<void> {
    return Promise.resolve();
  }
  detach(): void {}
  async *ask(): AsyncGenerator<SseEvent> {
    yield { type: 'done' };
  }
  apply(
    kind: ActuationRequest['kind'],
    _params: ActuationParams,
    changeId: ChangeId,
  ): Promise<ActuationResult> {
    return Promise.resolve({ ok: true, changeId, kind, location: 'A1' });
  }
  async *runCommands(
    task: string,
    _opts?: RunCommandsOptions,
  ): AsyncGenerator<SseEvent | CommandLoopEvent> {
    this.runTasks.push(task);
    yield { type: 'done', turn: 1, answer: '' };
  }
  plan(): Promise<{ plan: null; errors: string[]; needsClarification: boolean }> {
    return Promise.resolve({ plan: null, errors: [], needsClarification: false });
  }
  ingest(_e: HostEvent): Promise<void> {
    return Promise.resolve();
  }
}

const lister = (refs: ContextRef[]): ContextLister => ({
  listContext: () => Promise.resolve(refs),
});

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement | undefined;
let root: Root | undefined;

afterEach(() => {
  if (root) act(() => root!.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

function mount(
  controller: PanelController,
  surface: 'word' | 'excel' | 'powerpoint',
): HTMLDivElement {
  const el = document.createElement('div');
  container = el;
  document.body.appendChild(el);
  const r = createRoot(el);
  root = r;
  act(() => {
    r.render(createElement(App, { controller, surface }));
  });
  return el;
}

function setValue(input: HTMLInputElement, value: string): void {
  const setNative = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype,
    'value',
  )?.set;
  act(() => {
    setNative?.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

describe('<App/> parameterized quick action (Workstream H)', () => {
  it('clicking a parameterized chip opens the fill form instead of dispatching', () => {
    const assist = new RecordingAssist();
    const controller = new PanelController(assist, lister([]));
    const runSpy = vi.spyOn(controller, 'runCommands');
    const el = mount(controller, 'excel');

    const chip = el.querySelector('button[data-action-id="write-formula"]') as HTMLButtonElement;
    expect(chip).toBeTruthy();
    act(() => chip.dispatchEvent(new MouseEvent('click', { bubbles: true })));

    // The form is shown; nothing dispatched yet.
    expect(el.querySelector('[data-testid="quick-action-param-form"]')).toBeTruthy();
    expect(el.querySelector('[data-testid="qa-param-goal"]')).toBeTruthy();
    expect(runSpy).not.toHaveBeenCalled();
  });

  it('submit is disabled until the required value is filled (require-values-before-dispatch)', () => {
    const assist = new RecordingAssist();
    const controller = new PanelController(assist, lister([]));
    const el = mount(controller, 'excel');

    const chip = el.querySelector('button[data-action-id="write-formula"]') as HTMLButtonElement;
    act(() => chip.dispatchEvent(new MouseEvent('click', { bubbles: true })));

    const submit = el.querySelector(
      '[data-testid="quick-action-param-submit"]',
    ) as HTMLButtonElement;
    expect(submit.disabled).toBe(true);

    setValue(el.querySelector('[data-testid="qa-param-goal"]') as HTMLInputElement, 'sum of B:B');
    expect(submit.disabled).toBe(false);
  });

  it('filling and submitting dispatches the substituted seed — no literal {{…}} reaches the model', async () => {
    const assist = new RecordingAssist();
    const controller = new PanelController(assist, lister([]));
    const el = mount(controller, 'powerpoint');

    const chip = el.querySelector('button[data-action-id="draft-section"]') as HTMLButtonElement;
    expect(chip).toBeTruthy();
    act(() => chip.dispatchEvent(new MouseEvent('click', { bubbles: true })));

    setValue(el.querySelector('[data-testid="qa-param-topic"]') as HTMLInputElement, 'Q3 GTM');
    const form = el.querySelector('[data-testid="quick-action-param-form"]') as HTMLFormElement;
    await act(async () => {
      form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    });

    expect(assist.runTasks).toHaveLength(1);
    const task = assist.runTasks[0] ?? '';
    expect(task).toContain('Q3 GTM');
    expect(task).not.toContain('{{');
    // The form closes after dispatch.
    expect(el.querySelector('[data-testid="quick-action-param-form"]')).toBeFalsy();
  });
});
