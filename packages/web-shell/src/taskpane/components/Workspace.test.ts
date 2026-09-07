// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { quickActionsForSurface } from '@ge/contracts';
import { Composer, parseComposerInput } from './Composer.js';
import { ActionLibrary, filterActions } from './ActionLibrary.js';
import { invocationToSeed } from './quick-action-seed.js';
import { invocationToGrounding } from './App.js';
import { ContextStrip } from './ContextStrip.js';
import { MessageThread } from './MessageThread.js';
import { makeDemoController } from '../preview-interactive.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
let root: Root;
let container: HTMLDivElement;
function mount(element: JSX.Element): void {
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
  act(() => root.render(element));
}
function change(selector: string, value: string): void {
  const element = container.querySelector(selector) as
    | HTMLInputElement
    | HTMLSelectElement
    | HTMLTextAreaElement;
  const prototype =
    element instanceof HTMLTextAreaElement
      ? HTMLTextAreaElement.prototype
      : element instanceof HTMLSelectElement
        ? HTMLSelectElement.prototype
        : HTMLInputElement.prototype;
  act(() => {
    Object.getOwnPropertyDescriptor(prototype, 'value')?.set?.call(element, value);
    element.dispatchEvent(
      new Event(element instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }),
    );
  });
}
function click(selector: string): void {
  act(() => container.querySelector<HTMLButtonElement>(selector)!.click());
}
function submit(): void {
  act(() =>
    container
      .querySelector('form')!
      .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
  );
}
afterEach(() => {
  if (root) act(() => root.unmount());
  container?.remove();
  localStorage.clear();
  vi.restoreAllMocks();
});

function mountComposer(extra = {}): ReturnType<typeof vi.fn> {
  const onInvoke = vi.fn();
  mount(
    createElement(Composer, {
      surface: 'word',
      busy: false,
      onSend: vi.fn(),
      onCancel: vi.fn(),
      onInvoke,
      ...extra,
    }),
  );
  return onInvoke;
}

describe('structured workspace interactions', () => {
  it('routes picked intent and output controls into the invocation, preserving the original request', () => {
    const invoke = mountComposer();
    change('[aria-label="Task intent"]', 'rewrite');
    change('[aria-label="Response format"]', 'Decision brief');
    change('[aria-label="Writing style"]', 'Executive');
    change('textarea', 'Preserve the 99.9% commitment.');
    submit();
    expect(invoke).toHaveBeenCalledOnce();
    expect(invocationToSeed(invoke.mock.calls[0]![0])).toMatch(/^\/rewrite Preserve/);
    expect(invoke.mock.calls[0]![0]).toMatchObject({
      intent: 'rewrite',
      raw: 'Preserve the 99.9% commitment.',
      scope: { kind: 'selection' },
    });
    expect(invocationToSeed(invoke.mock.calls[0]![0])).toContain(
      'Output format: Decision brief. Writing style: Executive.',
    );
  });
  it('carries a source picker id into real request grounding, then clears per-request picks', () => {
    const id = 'projects/p/locations/global/collections/default_collection/dataStores/policy';
    const invoke = mountComposer({
      mentionOptions: { datastore: [{ id, label: 'Policy library' }] },
    });
    click('.composer-source-results button');
    expect(container.querySelector('.composer-picked')?.textContent).toContain('Policy library');
    change('textarea', 'What changed?');
    submit();
    const invocation = invoke.mock.calls[0]![0];
    expect(invocation.raw).toBe('What changed?');
    expect(invocationToGrounding(invocation)?.dataStoreSpecs).toEqual([{ dataStore: id }]);
    expect(container.querySelector('.composer-picked')).toBeNull();
  });
  it('removes a picked source from grounding before submission', () => {
    const invoke = mountComposer({
      mentionOptions: { datastore: [{ id: 'store-1', label: 'Policies' }] },
    });
    click('.composer-source-results button');
    click('.composer-picked button');
    change('textarea', 'Question');
    submit();
    expect(invoke.mock.calls[0]![0].mentions).toEqual([]);
  });
  it('typed verbs override the intent control and an unsupported verb cannot inherit a write', () => {
    const invoke = mountComposer();
    change('[aria-label="Task intent"]', 'rewrite');
    change('textarea', '/summarize the document');
    submit();
    expect(invoke.mock.calls[0]![0].intent).toBe('summarize');
    change('textarea', '/draft slides');
    submit();
    expect(invoke.mock.calls[1]![0].intent).toBeUndefined();
  });
  it('leaves pasted programs byte-for-byte intact despite response preferences', () => {
    const invoke = mountComposer();
    change('[aria-label="Response format"]', 'Comparison table');
    change('textarea', 'suggest "before" => "after"');
    submit();
    expect(invoke.mock.calls[0]![0].instruction).toBe('suggest "before" => "after"');
  });
  it('does not interpret an email address as a source mention', () => {
    expect(
      parseComposerInput('Contact user@unit.example about @this', { kind: 'selection' }).mentions,
    ).toEqual([{ kind: 'this' }]);
  });
  it('does not send while an IME is composing', () => {
    const invoke = mountComposer();
    change('textarea', '日本語');
    act(() =>
      container
        .querySelector('textarea')!
        .dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true }),
        ),
    );
    expect(invoke).not.toHaveBeenCalled();
  });
  it('does not offer unsupported actions even if their ids were previously pinned', () => {
    localStorage.setItem('ge.action-pins.v1', JSON.stringify(['tighten', 'create-chart']));
    mount(
      createElement(ActionLibrary, {
        surface: 'word',
        allowedIntents: ['ask'],
        disabled: false,
        onAction: vi.fn(),
      }),
    );
    expect(container.querySelector('[data-action-id="tighten"]')).toBeNull();
    expect(container.querySelector('[data-action-id="create-chart"]')).toBeNull();
    expect(container.querySelectorAll('[data-action-id]').length).toBeGreaterThan(0);
  });
  it('searches across task and outcome words without requiring slash syntax', () => {
    const actions = quickActionsForSurface('excel');
    expect(
      filterActions(actions, 'forecast variances', 'all', []).map((action) => action.id),
    ).toEqual(['excel-reconcile']);
    expect(
      filterActions(actions, '', 'pinned', ['excel-chart-brief']).map((action) => action.id),
    ).toEqual(['excel-chart-brief']);
  });
  it('survives disabled browser storage and prevents execution while busy', () => {
    vi.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('Unavailable');
    });
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('Unavailable');
    });
    const run = vi.fn();
    mount(createElement(ActionLibrary, { surface: 'word', disabled: true, onAction: run }));
    click('.pin-action');
    click('.library-run');
    expect(run).not.toHaveBeenCalled();
    expect(container.querySelector('.pin-action')?.getAttribute('aria-pressed')).toBe('true');
  });
  it('locks attachment changes during an in-flight turn or approval', () => {
    const toggle = vi.fn();
    mount(
      createElement(ContextStrip, {
        chips: [{ id: 'd', kind: 'document', title: 'Brief', attached: true }],
        disabled: true,
        onToggle: toggle,
        onReveal: vi.fn(),
        onRefresh: vi.fn(),
      }),
    );
    click('.smart-chip-remove');
    expect(toggle).not.toHaveBeenCalled();
  });
  it('never offers insertion or follow-up actions for incomplete streamed content', () => {
    mount(
      createElement(MessageThread, {
        messages: [
          {
            id: 'a',
            role: 'assistant',
            text: '| A | B |\n| --- | --- |\n| 1 | 2 |',
            streaming: true,
          },
        ],
        onInsertArtifact: vi.fn(),
        onFollowUp: vi.fn(),
      }),
    );
    expect(container.querySelector('.md-artifact-insert')).toBeNull();
    expect(container.querySelector('.answer-actions')).toBeNull();
  });
  it('allows answer follow-ups without an Excel insertion destination', () => {
    const follow = vi.fn();
    mount(
      createElement(MessageThread, {
        surface: 'excel',
        messages: [{ id: 'a', role: 'assistant', text: 'Analysis complete.' }],
        onFollowUp: follow,
        insertArtifactDisabledReason: 'Select a destination range first.',
      }),
    );
    const button = [
      ...container.querySelectorAll<HTMLButtonElement>('.answer-actions button'),
    ].find((b) => b.textContent === 'Decision brief')!;
    act(() => button.click());
    expect(follow).toHaveBeenCalledOnce();
  });
  it('uses a live controller in the demo and suspends execution at the real approval coordinator', async () => {
    const controller = makeDemoController('word');
    await controller.refreshContext();
    const id = controller.getState().chips[0]!.id;
    await controller.attach(id);
    expect(controller.getState().chips[0]!.attached).toBe(true);
    const run = controller.runCommands('Rewrite the selection');
    for (let i = 0; i < 20 && !controller.getState().pendingPlan; i++) await Promise.resolve();
    expect(controller.getState().pendingPlan?.effects[0]?.request.surface).toBe('word');
    expect(controller.getState().busy).toBe(true);
    controller.rejectPlan();
    await run;
    expect(controller.getState().pendingPlan).toBeUndefined();
    expect(controller.getState().busy).toBe(false);
    expect(controller.getState().messages.at(-1)?.text).toContain('rejected');
  });
});
