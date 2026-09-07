// @vitest-environment jsdom
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { ArtifactSummary } from '@ge/contracts';
import { getWorkflowRecipe } from '@ge/runtime';
import { WorkflowWorkbench } from './WorkflowWorkbench.js';
import { PanelController, type AssistLike, type PanelState } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
let root: Root | undefined;
let container: HTMLDivElement;
beforeEach(() => localStorage.clear());
afterEach(() => {
  if (root) act(() => root?.unmount());
  root = undefined;
  container?.remove();
});

const result: ArtifactSummary = {
  id: 'a_111111111111111111111111',
  hash: `sha256:${'1'.repeat(64)}`,
  title: 'Duplicate keys',
  createdAt: '2026-09-07T12:00:00.000Z',
  columns: [
    { name: 'c0', label: 'Key', type: 'string' },
    { name: 'c1', label: 'Occurrences', type: 'number' },
  ],
  sources: [
    {
      surface: 'excel',
      documentId: 'doc',
      locator: 'Orders!A1:B50',
      hash: `sha256:${'2'.repeat(64)}`,
    },
  ],
  lineage: { operation: 'query', parents: [] },
  rowCount: 2,
  preview: [
    ['ORD-1', 3],
    ['ORD-2', 2],
  ],
  truncated: false,
};

function setup(options: { disabled?: boolean; analysis?: boolean } = {}) {
  const base = new PanelController({} as AssistLike, { listContext: async () => [] });
  let state: PanelState = {
    ...base.getState(),
    workflowRecipesAvailable: true,
    ...(options.analysis !== false ? { analysis: { artifacts: [], offers: [] } } : {}),
  };
  const preview = vi.fn(async (_id: string, _inputs: Record<string, unknown>): Promise<void> => {});
  const write = vi.fn(async (): Promise<void> => {});
  const cancel = vi.fn();
  const controller = {
    runWorkflowRecipe: preview,
    runAnalysis: write,
    cancel,
    getState: () => state,
  } as unknown as PanelController;
  container = document.createElement('div');
  container.style.width = '320px';
  document.body.append(container);
  root = createRoot(container);
  const render = (patch: Partial<PanelState> = {}) => {
    state = { ...state, ...patch };
    act(() =>
      root!.render(
        createElement(WorkflowWorkbench, {
          state,
          controller,
          disabled: options.disabled ?? false,
        }),
      ),
    );
  };
  render();
  const ready = (
    artifact = result,
    writeState?: NonNullable<PanelState['workflowRun']>['write'],
  ) => {
    const recipeId = 'duplicate-rows';
    render({
      workflowRun: {
        runId: 'run-1',
        recipeId,
        recipeVersion: 1,
        status: 'ready',
        resultId: artifact.id,
        inputs: getWorkflowRecipe(recipeId).inputSchema.parse({ sourceRange: 'Orders!A1:B50' }),
        ...(writeState ? { write: writeState } : {}),
      },
      analysis: {
        artifacts: [
          artifact,
          { ...artifact, id: 'a_222222222222222222222222', title: 'Unrelated selected table' },
        ],
        selected: 'a_222222222222222222222222',
        offers: [],
      },
    });
  };
  return { preview, write, cancel, render, ready };
}

function button(name: string): HTMLButtonElement {
  const found = [...container.querySelectorAll('button')].find(
    (node) => node.textContent?.trim() === name,
  );
  if (!found) throw new Error(`Missing button: ${name}`);
  return found;
}

function choose(name = 'Find duplicate keys'): void {
  const choice = [...container.querySelectorAll<HTMLButtonElement>('.workflow-choice')].find(
    (node) => node.querySelector('strong')?.textContent === name,
  )!;
  act(() => choice.click());
}

function fill(name: string, value: string): void {
  const input = container.querySelector<HTMLInputElement>(`input[name="${name}"]`)!;
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

async function submit(selector = 'form[aria-label]'): Promise<void> {
  await act(async () => {
    container
      .querySelector(selector)!
      .dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
  });
}

describe('rendered workflow workbench', () => {
  it('offers three contract-backed workflows only when versioned analysis is available', () => {
    const { render } = setup();
    expect(container.querySelectorAll('.workflow-choice')).toHaveLength(3);
    expect(container.textContent).not.toMatch(/CLI|SDK|WASM|sessionless/);
    render({ analysis: undefined });
    expect(container.textContent).toBe('');
    render({ analysis: { artifacts: [], offers: [] }, workflowRecipesAvailable: false });
    expect(container.textContent).toBe('');
  });

  it('focuses the first source, validates before execution and converts visible column numbers once', async () => {
    const { preview } = setup();
    choose();
    expect(document.activeElement?.getAttribute('name')).toBe('sourceRange');
    await submit();
    expect(preview).not.toHaveBeenCalled();
    expect(container.querySelector('[role="alert"]')).not.toBeNull();
    expect(document.activeElement?.getAttribute('name')).toBe('sourceRange');
    fill('sourceRange', "'Current Orders'!A1:B50");
    fill('keyColumn', '2');
    await submit();
    expect(preview).toHaveBeenCalledWith('duplicate-rows', {
      sourceRange: "'Current Orders'!A1:B50",
      keyColumn: 1,
      headers: true,
      caseSensitive: true,
    });
    expect(preview.mock.calls[0]![1]).not.toHaveProperty('destination');
  });

  it('previews actual results and writes that result, independent of the globally selected table', async () => {
    const { preview, write, ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    await submit();
    expect(preview).toHaveBeenCalledTimes(1);
    ready();
    expect(container.querySelector('.workflow-sources')?.textContent).toContain('Orders!A1:B50');
    expect(container.querySelector('th')?.textContent).toBe('Key');
    expect(container.querySelector('.workflow-result-count')?.textContent).toContain(
      '2result rows',
    );
    expect(button('Review write').disabled).toBe(true);
    fill('destination', 'Results!A1');
    await submit('.workflow-write');
    expect(write).toHaveBeenCalledWith({
      kind: 'materialize',
      id: result.id,
      destination: 'Results!A1',
    });
  });

  it('invalidates the preview when source settings change and never targets a stale result', () => {
    const { ready, write } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready();
    fill('destination', 'Results!A1');
    fill('sourceRange', 'Next month!A1:B50');
    expect(container.querySelector('.workflow-result')).toBeNull();
    expect(container.textContent).toContain('Settings changed. Preview again');
    expect(write).not.toHaveBeenCalled();
  });

  it('keeps an empty result useful and does not offer an impossible write', () => {
    const { ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready({ ...result, rowCount: 0, preview: [] });
    expect(container.textContent).toContain('No duplicate nonblank keys found');
    expect(container.querySelector('.workflow-write')).toBeNull();
  });

  it('labels partial results and blocks writing incomplete data', () => {
    const { ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready({ ...result, truncated: true });
    expect(container.textContent).toContain('Partial result');
    expect(container.textContent).toContain('Use a smaller source range');
    expect(container.querySelector('.workflow-write')).toBeNull();
  });

  it.each(['written', 'uncertain', 'pending'] as const)(
    'suppresses another %s write to the same destination',
    (status) => {
      const { ready } = setup();
      choose();
      fill('sourceRange', 'Orders!A1:B50');
      ready(result, { destination: 'Results!A1', status });
      fill('destination', 'Results!A1');
      expect(
        (container.querySelector('.workflow-write button') as HTMLButtonElement).disabled,
      ).toBe(true);
      if (status === 'uncertain') expect(container.textContent).toContain('Check Recovery & undo');
    },
  );

  it('allows a rejected write to be reviewed again without presenting it as completed', () => {
    const { ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready(result, { destination: 'Results!A1', status: 'rejected' });
    fill('destination', 'Results!A1');
    expect(button('Review write').disabled).toBe(false);
    expect(container.textContent).toContain('Write declined');
    expect(container.textContent).not.toContain('Written and verified');
  });

  it('keeps a computed result inspectable when the document cannot accept writes', () => {
    const { ready, render } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready();
    render({ workflowWritesAvailable: false });
    expect(container.querySelector('.workflow-table')).not.toBeNull();
    expect(container.querySelector('.workflow-write')).toBeNull();
    expect(container.textContent).toContain('Writing is unavailable in this document');
  });

  it('requires recovery after an uncertain outcome even if the destination changes', () => {
    const { ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready(result, { destination: 'Results!A1', status: 'uncertain' });
    fill('destination', 'Other!A1');
    expect(button('Review write').disabled).toBe(true);
    expect(container.textContent).toContain('Check Recovery & undo');
  });

  it('requires a refreshed preview after the source changes before a write', async () => {
    const { ready, write, render } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    ready();
    fill('destination', 'Results!A1');
    write.mockImplementationOnce(async () => {
      render({ error: 'Source changed: Orders!A1:B50. Capture and analyze it again.' });
    });
    await submit('.workflow-write');
    expect(container.textContent).toContain('A source has changed. Refresh the preview');
    expect(button('Review write').disabled).toBe(true);
    expect(button('Refresh preview').disabled).toBe(false);
    // A controller ownership race that declines to start a new run must not clear stale evidence.
    await submit();
    expect(button('Review write').disabled).toBe(true);
  });

  it('gates stale asynchronous errors after switching workflows and prevents overlapping runs', async () => {
    const { preview, cancel } = setup();
    let reject: (error: Error) => void = () => {};
    preview.mockImplementationOnce(
      () =>
        new Promise((_, fail) => {
          reject = fail;
        }),
    );
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    await submit();
    expect(button('Preparing preview…').disabled).toBe(true);
    act(() => button('Cancel').click());
    expect(cancel).toHaveBeenCalledTimes(1);
    choose('Summarize amounts');
    await act(async () => reject(new Error('Old workflow failed')));
    expect(container.textContent).not.toContain('Old workflow failed');
    expect(container.querySelector('form')?.getAttribute('aria-label')).toBe(
      'Summarize amounts settings',
    );
    expect(preview).toHaveBeenCalledTimes(1);
  });

  it('shows a failed matching run, focuses its error and preserves settings for retry', () => {
    const { render } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    render({
      workflowRun: {
        runId: 'failed',
        recipeId: 'duplicate-rows',
        recipeVersion: 1,
        status: 'failed',
        inputs: getWorkflowRecipe('duplicate-rows').inputSchema.parse({
          sourceRange: 'Orders!A1:B50',
        }),
        error: 'Source changed: Orders!A1:B50. Capture and analyze it again.',
      },
    });
    expect(container.textContent).toContain('Preview could not finish');
    expect(document.activeElement?.getAttribute('role')).toBe('alert');
    expect(button('Preview result').disabled).toBe(false);
  });

  it('saves only on explicit action and restores parameters without running or granting approval', () => {
    const { preview, write, ready } = setup();
    choose();
    fill('sourceRange', 'Orders!A1:B50');
    expect(localStorage.length).toBe(0);
    ready();
    fill('destination', 'Results!A1');
    act(() => button('Save settings').click());
    const stored = localStorage.getItem(localStorage.key(0)!)!;
    expect(stored).toContain('Orders!A1:B50');
    expect(stored).not.toContain(result.id);
    expect(stored).not.toContain('approval');
    choose('Summarize amounts');
    choose();
    act(() => button('Use saved settings').click());
    expect((container.querySelector('input[name="sourceRange"]') as HTMLInputElement).value).toBe(
      'Orders!A1:B50',
    );
    expect(preview).not.toHaveBeenCalled();
    expect(write).not.toHaveBeenCalled();
  });

  it('disables execution and navigation while another controller task owns the panel', () => {
    setup({ disabled: true });
    expect([...container.querySelectorAll('button')].every((node) => node.disabled)).toBe(true);
  });
});
