// @vitest-environment jsdom
import { openPlanEffects } from '../test-harness/mount-stack.js';
import { describe, it, expect, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { App, type AppProps } from './components/App.js';
import { makeMockController } from './preview.js';
import { FIXTURE_STATE } from './preview-fixtures.js';
import type { Surface } from '@ge/contracts';

/**
 * Render smoke test for the task pane: mount the real <App/> over the fake controller and the full
 * preview fixtures, then assert every card actually renders — thread (with citations), run-steps,
 * the plan-approval card with its effect-set, the write-approval card with its verbatim command,
 * proposals, the error banner, and the busy state. This both proves the preview harness wires up
 * and guards the UI against regressions. It exercises the view only; no controller/approval logic.
 */

// Tell React this is an act-aware environment so state updates flush deterministically.
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(surface: Surface = 'word', props: Partial<AppProps> = {}): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  const controller = makeMockController(FIXTURE_STATE);
  act(() => {
    root.render(createElement(App, { controller, surface, ...props }));
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('<App/> render smoke', () => {
  it('renders the header and marks the panel busy for the chosen surface', () => {
    render();
    expect(container.textContent).toContain('Gemini Enterprise');
    const panel = container.querySelector('.panel');
    expect(panel?.getAttribute('data-surface')).toBe('word');
    expect(panel?.getAttribute('aria-busy')).toBe('true');
  });

  it('keeps the sidebar destinations labeled at narrow-pane density', () => {
    render();
    const labels = [...container.querySelectorAll('.tw-rail .tw-label')].map(
      (node) => node.textContent,
    );
    expect(labels).toEqual(expect.arrayContaining(['Chat', 'Tools', 'History']));
  });

  it('collapses the sidebar to a menu button and opens it again', () => {
    localStorage.clear();
    render();
    act(() =>
      container.querySelector<HTMLButtonElement>('[aria-label="Collapse sidebar"]')?.click(),
    );
    expect(container.querySelector('.tw-rail')).toBeNull();
    expect(localStorage.getItem('ge.sidebar')).toBe('closed');
    act(() => container.querySelector<HTMLButtonElement>('[aria-label="Open sidebar"]')?.click());
    expect(container.querySelector('.tw-rail')).not.toBeNull();
    localStorage.clear();
  });

  it('opens Settings with the answer options and no skills catalog', () => {
    render('word');
    act(() => container.querySelector<HTMLButtonElement>('[data-panel="settings"]')?.click());
    expect(container.querySelector('.tw-modal-title')?.textContent).toContain('Settings');
    expect(container.querySelector('[aria-label="Response format"]')).not.toBeNull();
    expect(container.querySelector('#ge-catalog-settings')).toBeNull();
    expect(container.querySelector('#tw-panel-settings')?.textContent).not.toMatch(/skill/i);
  });

  it('renders the conversation thread with an assistant message and its citations', () => {
    render();
    const thread = container.querySelector('.thread[role="log"]');
    expect(thread).not.toBeNull();
    expect(thread?.textContent).toContain('Northwind Cloud');
    // Citation pills from the streamed sources.
    expect(container.querySelectorAll('.cite').length).toBeGreaterThan(0);
    expect(container.querySelector('.cites')?.getAttribute('aria-label')).toBe('Citations');
  });

  it('renders the loader, not the caret, on the in-flight assistant message while it works', () => {
    render();
    expect(container.querySelector('.caret')).toBeNull();
    const activity = container.querySelector('.message-activity[role="status"]');
    expect(activity?.textContent).toContain('Checking the selected policy');
  });

  it('renders the sources row, which opens to list every chip', () => {
    render();
    const toggle = container.querySelector<HTMLButtonElement>('.context-strip-toggle');
    expect(toggle?.getAttribute('aria-expanded')).toBe('false');
    act(() => toggle?.click());
    const add = [
      ...container.querySelectorAll<HTMLButtonElement>('.context-strip .text-control'),
    ].find((b) => /^(Add|Manage) sources/.test(b.textContent ?? ''));
    act(() => add?.click());
    expect(container.querySelectorAll('.smart-chip').length).toBe(FIXTURE_STATE.chips.length);
  });

  it('opens the searchable action library and disables execution at the approval gate', () => {
    render();
    act(() => container.querySelector<HTMLButtonElement>('[data-panel="actions"]')?.click());
    expect(container.querySelector('.tw-modal-title')?.textContent).toContain('Tools');
    expect(container.querySelector('.action-library input[type="search"]')).not.toBeNull();
    const buttons = [...container.querySelectorAll<HTMLButtonElement>('.library-run')];
    expect(buttons.length).toBeGreaterThan(0);
    expect(buttons.every((button) => button.disabled)).toBe(true);
  });

  it('starts with the sidebar collapsed in a narrow pane, unless the user chose otherwise', () => {
    const width = window.innerWidth;
    localStorage.clear();
    Object.defineProperty(window, 'innerWidth', { configurable: true, value: 320 });
    try {
      render();
      expect(container.querySelector('.tw-rail')).toBeNull();
      expect(container.querySelector('[aria-label="Open sidebar"]')).not.toBeNull();
      act(() => root.unmount());
      container.remove();
      localStorage.setItem('ge.sidebar', 'open');
      render();
      expect(container.querySelector('.tw-rail')).not.toBeNull();
    } finally {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: width });
      localStorage.clear();
    }
  });

  it('follows the thread to the newest message unless the user scrolled up to read', () => {
    const base = { ...FIXTURE_STATE, busy: false };
    const withMessages = (extra: typeof base.messages) => ({
      ...base,
      messages: [...base.messages, ...extra],
    });
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
    const show = (state: typeof base): void =>
      act(() => {
        root.render(createElement(App, { controller: makeMockController(state), surface: 'word' }));
      });
    show(base);
    const region = container.querySelector<HTMLElement>('.thread-region')!;
    Object.defineProperty(region, 'scrollHeight', { configurable: true, value: 1000 });
    Object.defineProperty(region, 'clientHeight', { configurable: true, value: 200 });

    // A new question jumps to the bottom.
    show(withMessages([{ id: 'u2', role: 'user', text: 'Next question' }]));
    expect(region.scrollTop).toBe(1000);

    // The user scrolls up to read: a streaming answer does not pull them down.
    region.scrollTop = 100;
    act(() => region.dispatchEvent(new Event('scroll')));
    show(
      withMessages([
        { id: 'u2', role: 'user', text: 'Next question' },
        { id: 'a2', role: 'assistant', text: 'Part of the answer' },
      ]),
    );
    expect(region.scrollTop).toBe(100);
  });

  it('closes an open panel when a decision card appears behind it', () => {
    const idle = {
      ...FIXTURE_STATE,
      busy: false,
      pendingPlan: undefined,
      pendingWrite: undefined,
      pendingCommandPlan: undefined,
      pendingShare: undefined,
    };
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
    act(() => {
      root.render(createElement(App, { controller: makeMockController(idle), surface: 'excel' }));
    });
    act(() => container.querySelector<HTMLButtonElement>('[data-panel="actions"]')?.click());
    expect(container.querySelector('.tw-modal')?.hasAttribute('hidden')).toBe(false);
    // e.g. a workflow run from Tools stages its write: the plan approval shows in the main column.
    act(() => {
      root.render(
        createElement(App, { controller: makeMockController(FIXTURE_STATE), surface: 'excel' }),
      );
    });
    expect(container.querySelector('.plan-approval')).not.toBeNull();
    expect(container.querySelector('.tw-modal')?.hasAttribute('hidden')).toBe(true);
  });

  it('offers each available action once in the library for Excel', () => {
    render('excel');
    act(() => container.querySelector<HTMLButtonElement>('[data-panel="actions"]')?.click());
    for (const id of ['create-chart', 'summarize-range', 'find-anomalies', 'excel-reconcile']) {
      expect(container.querySelectorAll(`.action-library [data-action-id="${id}"]`)).toHaveLength(
        1,
      );
    }
  });

  it('renders suggestions', () => {
    render();
    expect(container.querySelector('.suggestions[aria-label="Suggestions"]')).not.toBeNull();
    expect(container.querySelector('.suggestion')?.textContent).toContain('renewal-risk');
  });

  it('renders the run-steps transcript as a polite live region', () => {
    render();
    const steps = container.querySelector('.run-steps[aria-label="Command loop steps"]');
    expect(steps).not.toBeNull();
    expect(steps?.getAttribute('aria-live')).toBe('polite');
    const toggle = steps?.querySelector<HTMLButtonElement>('.run-steps-toggle');
    const list = steps?.querySelector<HTMLOListElement>('.run-steps-list');
    expect(toggle?.getAttribute('aria-expanded')).toBe('false');
    expect(list?.hidden).toBe(true);
    expect(toggle?.textContent).toContain(`${FIXTURE_STATE.steps.length} steps`);
    expect(steps?.querySelectorAll('.run-step').length).toBe(FIXTURE_STATE.steps.length);
    act(() => toggle?.click());
    expect(toggle?.getAttribute('aria-expanded')).toBe('true');
    expect(list?.hidden).toBe(false);
  });

  it('renders the plan-approval card with the full effect-set and a summary', () => {
    render();
    const plan = container.querySelector('.plan-approval[aria-label="Plan approval required"]');
    expect(plan).not.toBeNull();
    expect(plan?.querySelector('.plan-approval-title')?.textContent).toBe('Review 5 changes');
    expect(plan?.querySelectorAll('.plan-effect').length).toBe(
      FIXTURE_STATE.pendingPlan?.effects.length,
    );
    // Each effect renders its verbatim command line.
    openPlanEffects(container);
    expect(plan?.textContent).toContain('set Sales!F2 =C2-D2');
    expect(plan?.textContent).toContain('comment Sales!F2');
    // The approve/reject actions gate real writes — both must be reachable buttons.
    const buttons = plan?.querySelectorAll('.act button') ?? [];
    expect(buttons.length).toBe(2);
  });

  it('renders the write-approval card with the verbatim command line', () => {
    render();
    const write = container.querySelector('.approval[aria-label="Write approval required"]');
    expect(write).not.toBeNull();
    const cmd = write?.querySelector('.cmd');
    expect(cmd?.textContent).toBe(FIXTURE_STATE.pendingWrite?.command);
    expect(write?.querySelectorAll('.act button').length).toBe(2);
  });

  it('renders proposal cards in their respective statuses', () => {
    render();
    const proposals = container.querySelector('.proposals[aria-label="Proposed changes"]');
    expect(proposals).not.toBeNull();
    expect(proposals?.querySelectorAll('.card').length).toBe(FIXTURE_STATE.proposals.length);
    expect(proposals?.querySelector('.status-applied')).not.toBeNull();
    expect(proposals?.querySelector('.status-degraded')).not.toBeNull();
    // The pending proposal exposes an Accept action.
    expect(proposals?.textContent).toContain('Accept change');
  });

  it('renders the error banner as an alert', () => {
    render();
    const err = container.querySelector('.panel-error[role="alert"]');
    expect(err).not.toBeNull();
    expect(err?.textContent).toContain(FIXTURE_STATE.error ?? '');
  });

  it('renders the composer with an accessible ask field', () => {
    render();
    const input = container.querySelector('textarea#ask');
    expect(input).not.toBeNull();
    // While busy, the send button flips to Cancel.
    expect(container.querySelector('.snd.cancel')).not.toBeNull();
  });

  it('has no Skills item in the sidebar, even when skills are registered', () => {
    render();
    expect(FIXTURE_STATE.skills?.length).toBeGreaterThan(0);
    expect(container.querySelector('[data-panel="skills"]')).toBeNull();
    expect(container.querySelector('#tw-panel-skills')).toBeNull();
    const labels = [...container.querySelectorAll('.tw-rail .tw-label')].map((n) => n.textContent);
    expect(labels).not.toContain('Skills');
  });

  it('opens the conversations sheet with the active session and continue controls', () => {
    render();
    const sessionsButton = container.querySelector<HTMLButtonElement>('[data-panel="sessions"]');
    act(() => sessionsButton?.click());
    const sessions = container.querySelector('.sessions[aria-label="Conversations"]');
    expect(sessions).not.toBeNull();
    expect(sessions?.textContent).toContain('Google SWE mock schedule');
    expect(sessions?.querySelector('.session-row.active')).not.toBeNull();
    expect(sessions?.querySelectorAll('.session-resume').length).toBeGreaterThan(0);
  });

  it('expands a plan effect to reveal its target and dry-run before→after preview', () => {
    render();
    const plan = container.querySelector('.plan-approval');
    // Each effect head is a collapsed, expandable button.
    const head = plan?.querySelector<HTMLButtonElement>('.plan-effect-head');
    expect(head?.getAttribute('aria-expanded')).toBe('false');
    act(() => head?.click());
    expect(head?.getAttribute('aria-expanded')).toBe('true');
    const detail = plan?.querySelector('.plan-effect-detail');
    expect(detail).not.toBeNull();
    // The first effect's dry-run resolves a before→after preview.
    expect(detail?.querySelector('.diff-after')?.textContent).toBe('$184,000');
    expect(detail?.textContent).toContain('Sales!F2');
  });

  it('opens a citation source-detail popover with title and link', () => {
    render();
    const cite = container.querySelector<HTMLButtonElement>('.cite-btn');
    expect(cite?.getAttribute('aria-expanded')).toBe('false');
    act(() => cite?.click());
    expect(cite?.getAttribute('aria-expanded')).toBe('true');
    const detail = container.querySelector('.cite-detail');
    expect(detail).not.toBeNull();
    expect(detail?.querySelector('.cite-detail-title')).not.toBeNull();
    expect(detail?.querySelector('a.cite-detail-link')).not.toBeNull();
  });

  it('renders the Excel formula-first body and entity card on a write-cells proposal', () => {
    render();
    const proposals = container.querySelector('.proposals');
    // Formula-first: the value renders as a code formula against its range target.
    const formula = proposals?.querySelector('.proposal-formula .formula');
    expect(formula?.textContent).toBe('=C2-D2');
    expect(proposals?.querySelector('.cell-target')?.textContent).toBe('Sales!F2');
    // The linked-entity card treatment.
    const entity = proposals?.querySelector('.entity-card');
    expect(entity).not.toBeNull();
    expect(entity?.querySelector('.entity-title')?.textContent).toBe('Northwind Cloud');
  });

  it('renders the Word redline body on a tracked-change proposal', () => {
    render();
    const redline = container.querySelector('.proposal-redline');
    expect(redline).not.toBeNull();
    expect(redline?.querySelector('del.redline-del')?.textContent).toBe('SLA of 99.5%');
    expect(redline?.querySelector('ins.redline-ins')?.textContent).toBe(
      'SLA of 99.9% (FSI standard)',
    );
  });

  it('drills into the provenance of an applied write', () => {
    render();
    const toggle = container.querySelector<HTMLButtonElement>('.prov-toggle');
    expect(toggle?.getAttribute('aria-expanded')).toBe('false');
    act(() => toggle?.click());
    expect(toggle?.getAttribute('aria-expanded')).toBe('true');
    const prov = container.querySelector('.provenance[aria-label="Change provenance"]');
    expect(prov).not.toBeNull();
    expect(prov?.textContent).toContain('contract-review-agent@v2');
    expect(prov?.textContent).toContain('v.k@acme.com');
    expect(prov?.querySelectorAll('.prov-sources li').length).toBe(2);
  });
});
