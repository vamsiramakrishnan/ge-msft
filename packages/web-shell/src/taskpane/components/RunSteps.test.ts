// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { RunSteps } from './RunSteps.js';
import type { RunStep } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(steps: RunStep[]): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(createElement(RunSteps, { steps }));
  });
}

afterEach(() => {
  if (root) act(() => root.unmount());
  container?.remove();
});

describe('RunSteps', () => {
  it('renders workspace artifacts as compact expandable cards', () => {
    render([
      {
        id: 's1',
        kind: 'read-result',
        text: 'grep schedule.tsv · 1 match',
        artifact: {
          title: 'ws:1 · schedule.tsv',
          meta: ['tsv · text/tab-separated-values', '20 lines · 1.1 KB'],
          matches: [{ line: 7, text: 'Deep Work' }],
        },
      },
    ]);

    expect(container.querySelector('.run-steps-latest')?.textContent).toContain(
      'grep schedule.tsv · 1 match',
    );
    expect(container.querySelector<HTMLOListElement>('.run-steps-list')?.hidden).toBe(true);

    const toggle = container.querySelector<HTMLButtonElement>('.run-steps-toggle');
    act(() => toggle?.click());

    expect(container.querySelector('.workspace-artifact-card')?.textContent).toContain(
      'ws:1 · schedule.tsv',
    );
    expect(container.querySelector('.workspace-artifact-line')?.textContent).toBe('L7');
  });
});
