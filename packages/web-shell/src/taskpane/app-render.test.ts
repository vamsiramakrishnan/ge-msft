// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { App } from './components/App.js';
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

function render(surface: Surface = 'word'): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  const controller = makeMockController(FIXTURE_STATE);
  act(() => {
    root.render(createElement(App, { controller, surface }));
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

  it('renders a surface-aware command center above the workstream', () => {
    render();
    const center = container.querySelector(
      '.surface-center[aria-label="Word workspace command center"]',
    );
    expect(center).not.toBeNull();
    expect(center?.querySelector('.surface-title')?.textContent).toBe('Word workspace');
    expect(center?.querySelectorAll('.surface-action').length).toBe(3);
    expect(center?.textContent).toContain('Gate');
  });

  it('renders surface-specific primary actions and keeps the secondary row deduplicated', () => {
    render('excel');
    const center = container.querySelector(
      '.surface-center[aria-label="Excel workspace command center"]',
    );
    expect(center?.querySelector('[data-action-id="create-chart"]')).not.toBeNull();
    expect(center?.querySelector('[data-action-id="summarize-range"]')).not.toBeNull();
    expect(center?.querySelector('[data-action-id="find-anomalies"]')).not.toBeNull();
    expect(container.querySelector('.quick-actions [data-action-id="create-chart"]')).toBeNull();
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
    const input = container.querySelector('textarea#ask');
    expect(input).not.toBeNull();
    // While busy, the send button flips to Cancel.
    expect(container.querySelector('.snd.cancel')).not.toBeNull();
  });

  it('renders the skills surface with each skill signature, def confirmation and an invoke action', () => {
    render();
    const skills = container.querySelector('.skills[aria-label="Skills"]');
    expect(skills).not.toBeNull();
    expect(skills?.querySelectorAll('.skill').length).toBe(FIXTURE_STATE.skills?.length);
    // The first skill renders its signature, registration badge, verbatim def and an Invoke button.
    expect(skills?.querySelector('.skill-sig')?.textContent).toContain('flag-vendor-risk');
    expect(skills?.querySelector('.skill-badge')?.textContent).toContain('registered');
    expect(skills?.querySelector('.skill-def')?.textContent).toContain('def flag-vendor-risk');
    expect(skills?.querySelector('[aria-label="Invoke flag-vendor-risk"]')).not.toBeNull();
    // Param fields are bindable inputs prefilled from the declared examples.
    const param = skills?.querySelector<HTMLInputElement>('input.skill-param-input');
    expect(param?.value).toBe('Northwind Cloud');
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
