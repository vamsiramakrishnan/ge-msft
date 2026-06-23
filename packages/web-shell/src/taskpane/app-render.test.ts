// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { App } from './components/App.js';
import { makeMockController } from './preview.js';
import { FIXTURE_STATE } from './preview-fixtures.js';

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

function render(): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  const controller = makeMockController(FIXTURE_STATE);
  act(() => {
    root.render(createElement(App, { controller, surface: 'word' }));
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

  it('renders the conversation thread with an assistant message and its citations', () => {
    render();
    const thread = container.querySelector('.thread[role="log"]');
    expect(thread).not.toBeNull();
    expect(thread?.textContent).toContain('Northwind Cloud');
    // Citation pills from the streamed sources.
    expect(container.querySelectorAll('.cite').length).toBeGreaterThan(0);
    expect(container.querySelector('.cites')?.getAttribute('aria-label')).toBe('Citations');
  });

  it('renders the streaming caret on the in-flight assistant message', () => {
    render();
    expect(container.querySelector('.caret')).not.toBeNull();
  });

  it('renders the context tray with attached and available chips', () => {
    render();
    const tray = container.querySelector('.unit[aria-label="Research unit grounding scope"]');
    expect(tray).not.toBeNull();
    expect(tray?.querySelector('[aria-label="Attached sources"]')).not.toBeNull();
    expect(tray?.querySelector('[aria-label="Available to attach"]')).not.toBeNull();
    expect(container.querySelectorAll('.chip').length).toBe(FIXTURE_STATE.chips.length);
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
    expect(steps?.querySelectorAll('.run-step').length).toBe(FIXTURE_STATE.steps.length);
  });

  it('renders the plan-approval card with the full effect-set and a summary', () => {
    render();
    const plan = container.querySelector('.plan-approval[aria-label="Plan approval required"]');
    expect(plan).not.toBeNull();
    expect(plan?.querySelector('.pin')?.textContent).toBe('3 writes + 2 comments');
    expect(plan?.querySelectorAll('.plan-effect').length).toBe(
      FIXTURE_STATE.pendingPlan?.effects.length,
    );
    // Each effect renders its verbatim command line.
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
    const input = container.querySelector('input#ask');
    expect(input).not.toBeNull();
    // While busy, the send button flips to Cancel.
    expect(container.querySelector('.snd.cancel')).not.toBeNull();
  });
});
