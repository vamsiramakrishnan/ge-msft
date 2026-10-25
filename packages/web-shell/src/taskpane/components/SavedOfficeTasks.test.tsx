// @vitest-environment jsdom
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PanelController, PanelState } from '../../controller.js';
import { PanelController as RealController, type AssistLike } from '../../controller.js';
import { SavedOfficeTasks } from './SavedOfficeTasks.js';
import { savedOfficeTaskPreview } from '../preview-saved-tasks.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
let root: Root | undefined;
let container: HTMLDivElement | undefined;
afterEach(() => {
  act(() => root?.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

function setup(mode = 'ready', disabled = false) {
  const base = new RealController({} as AssistLike, { listContext: async () => [] });
  let state: PanelState = {
    ...base.getState(),
    savedOfficeTasks: savedOfficeTaskPreview(new URLSearchParams(`saved-tasks=${mode}`), 'word'),
  };
  const methods = {
    refreshSavedOfficeTasks: vi.fn(async () => {}),
    inspectSavedOfficeTask: vi.fn(async () => {}),
    resumeSavedOfficeTask: vi.fn(async () => {}),
    archiveSavedOfficeTask: vi.fn(async () => {}),
    saveLiteralOfficeTask: vi.fn(async () => {}),
  };
  const controller = methods as unknown as PanelController;
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
  const render = (patch: Partial<PanelState> = {}) => {
    state = { ...state, ...patch };
    act(() => root!.render(createElement(SavedOfficeTasks, { state, controller, disabled })));
  };
  render();
  return { ...methods, render, state };
}
const button = (text: string): HTMLButtonElement => {
  const found = [...container!.querySelectorAll('button')].find(
    (element) => element.textContent === text,
  );
  if (!found) throw new Error(`Missing button ${text}`);
  return found;
};
function fill(selector: string, value: string) {
  const input = container!.querySelector<HTMLInputElement | HTMLTextAreaElement>(selector)!;
  act(() => {
    const descriptor = Object.getOwnPropertyDescriptor(
      input.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype,
      'value',
    )!;
    descriptor.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

describe('saved Office task pane', () => {
  it('never lists or resumes automatically; explicit inspection is a separate action', () => {
    const f = setup();
    expect(f.refreshSavedOfficeTasks).not.toHaveBeenCalled();
    expect(f.resumeSavedOfficeTask).not.toHaveBeenCalled();
    act(() => button('Inspect task').click());
    expect(f.inspectSavedOfficeTask).toHaveBeenCalledWith('month-close-memo');
    act(() => button('Review remaining').click());
    expect(f.resumeSavedOfficeTask).toHaveBeenCalledWith('month-close-memo');
    expect(container!.textContent).toContain('Nothing resumes automatically');
  });
  it('keeps unknown or unsupported work inspectable and blocks review/hiding', () => {
    setup('blocked');
    expect(button('Review remaining').disabled).toBe(true);
    expect(button('Hide completed task').disabled).toBe(true);
    expect(container!.textContent).toContain('unknown outcome');
    expect(container!.querySelector('[role="status"]')).not.toBeNull();
  });
  it('offers hiding only for completed tasks and never calls it undo or storage cleanup', () => {
    const f = setup('complete');
    expect(button('Review remaining').disabled).toBe(true);
    expect(button('Hide completed task').disabled).toBe(false);
    act(() => button('Hide completed task').click());
    expect(f.archiveSavedOfficeTask).toHaveBeenCalledWith('month-close-memo');
    expect(container!.textContent).toContain('safety history stays recorded');
    expect(container!.textContent).not.toContain('free storage');
  });
  it('requires a separate literal-input attestation and clears it when commands change', () => {
    const f = setup('empty');
    fill('input.field-control', 'draft-memo');
    fill('textarea', '/fill-content-control contentControlId="memo" text="Literal answer"');
    expect(button('Save and inspect').disabled).toBe(true);
    const confirm = container!.querySelector<HTMLInputElement>('input[type="checkbox"]')!;
    act(() => confirm.click());
    expect(button('Save and inspect').disabled).toBe(false);
    act(() =>
      container!
        .querySelector('form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    );
    expect(f.saveLiteralOfficeTask).toHaveBeenCalledWith(
      'draft-memo',
      '/fill-content-control contentControlId="memo" text="Literal answer"',
      true,
    );
    expect(f.resumeSavedOfficeTask).not.toHaveBeenCalled();
    fill('textarea', 'Changed instructions');
    expect(confirm.checked).toBe(false);
    expect(button('Save and inspect').disabled).toBe(true);
  });
  it('disables every mutating control while another turn or approval owns the pane', () => {
    const f = setup('ready', true);
    expect(
      [...container!.querySelectorAll('button,input,textarea')].every(
        (element) => (element as HTMLButtonElement).disabled,
      ),
    ).toBe(true);
    act(() => button('Review remaining').click());
    expect(f.resumeSavedOfficeTask).not.toHaveBeenCalled();
  });
  it('renders command source as text and hides the section when the runtime API is absent', () => {
    const f = setup();
    const selected = structuredClone(f.state.savedOfficeTasks!.selected!);
    selected.record.program.stages[0]!.program = '<img src=x onerror="window.bad=true">';
    f.render({ savedOfficeTasks: { ...f.state.savedOfficeTasks!, selected } });
    expect(container!.querySelector('img')).toBeNull();
    expect(container!.querySelector('pre')?.textContent).toContain('<img');
    f.render({ savedOfficeTasks: undefined });
    expect(container!.textContent).toBe('');
  });
});
