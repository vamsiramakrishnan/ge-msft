// @vitest-environment jsdom
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { AnalysisWorkbench } from './AnalysisWorkbench.js';
import { PanelController, type AssistLike, type PanelState } from '../../controller.js';
import { makeCellSnapshot } from '@ge/contracts';
import { AnalysisWorkspace } from '@ge/runtime';
import type { DocBridge } from '@ge/runtime';
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
let root: Root;
let container: HTMLDivElement;
afterEach(() => {
  if (root) act(() => root.unmount());
  container?.remove();
});
async function setup(disabled = false) {
  const workspace = new AnalysisWorkspace(
    {
      surface: 'excel',
      captureCells: async (locator: string) =>
        makeCellSnapshot({
          surface: 'excel',
          documentId: 'd',
          locator,
          values: [
            ['Invoice', 'Amount', 'Currency'],
            ['A', '12.00', 'USD'],
          ],
        }),
    } as DocBridge,
    async () => ({
      query: async () => ({ columns: [], rows: [], truncated: false, durationMs: 0 }),
      dispose() {},
    }),
  );
  await workspace.execute({ kind: 'capture', range: 'Invoices!A1:C2' });
  await workspace.execute({ kind: 'capture', range: 'Payments!A1:C2' });
  const controller = new PanelController({} as AssistLike, { listContext: async () => [] });
  const run = vi.spyOn(controller, 'runAnalysis').mockResolvedValue();
  const state: PanelState = { ...controller.getState(), analysis: workspace.state() };
  container = document.createElement('div');
  document.body.append(container);
  root = createRoot(container);
  act(() => root.render(createElement(AnalysisWorkbench, { state, controller, disabled })));
  return { run, state };
}
function fill(selector: string, value: string) {
  const el = container.querySelector(selector) as HTMLInputElement;
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(el, value);
    el.dispatchEvent(new Event('input', { bubbles: true }));
  });
}
describe('rendered data workbench', () => {
  it('renders inspectable tables, meaningful column labels and explicit destinations', async () => {
    const { run, state } = await setup();
    expect(container.textContent).toContain('2 tables');
    expect(container.querySelectorAll('.analysis-artifact')).toHaveLength(2);
    expect(container.querySelector('th')?.textContent).toBe('Invoice');
    expect(container.querySelector('[aria-label="Write destination"]')).not.toBeNull();
    fill('[aria-label="Write destination"]', 'Results!A1');
    act(() =>
      container
        .querySelector('.analysis-result form')!
        .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true })),
    );
    expect(run).toHaveBeenCalledWith({
      kind: 'materialize',
      id: state.analysis!.selected,
      destination: 'Results!A1',
    });
  });
  it('disables every action during task ownership or approval', async () => {
    await setup(true);
    expect(
      [...container.querySelectorAll('button,input,select,textarea')].every(
        (el) => (el as HTMLButtonElement).disabled,
      ),
    ).toBe(true);
  });
  it('validates currency and tolerance without emitting an invalid action', async () => {
    const { run } = await setup();
    const reconcile = container.querySelector('.analysis-primary') as HTMLButtonElement;
    expect(reconcile.disabled).toBe(false);
    fill('[aria-label="Reconciliation tolerance"]', '-1');
    expect(reconcile.disabled).toBe(true);
    fill('[aria-label="Reconciliation tolerance"]', '0.01');
    fill('[aria-label="Single currency"]', 'U');
    expect(reconcile.disabled).toBe(true);
    expect(run).not.toHaveBeenCalled();
  });
});
