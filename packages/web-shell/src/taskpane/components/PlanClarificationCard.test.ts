// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  clarificationChoices,
  PlanClarificationCard,
  type PlanClarificationCardProps,
} from './PlanClarificationCard.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(props: PlanClarificationCardProps): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(createElement(PlanClarificationCard, props));
  });
}

afterEach(() => {
  if (root) act(() => root.unmount());
  container?.remove();
});

describe('PlanClarificationCard', () => {
  it('maps common planner questions to suggested selector choices', () => {
    expect(clarificationChoices('What would you like the chart to show?')).toContain(
      'Hours per activity (whole week)',
    );
    expect(clarificationChoices('What routine should the schedule reflect?')).toContain(
      'Google SWE schedule',
    );
  });

  it('submits a suggested answer back to the planner flow', () => {
    const onAnswer = vi.fn();
    render({
      pending: {
        task: '/visualize insert a chart',
        questions: ['What would you like the chart to show?'],
      },
      onAnswer,
    });

    const choice = [...container.querySelectorAll<HTMLButtonElement>('.clarification-choice')].find(
      (button) => button.textContent === 'Hours per activity (whole week)',
    );
    expect(choice).not.toBeUndefined();
    act(() => choice?.click());
    expect(onAnswer).toHaveBeenCalledWith('Hours per activity (whole week)');
  });

  it('supports a custom answer when suggestions are not enough', () => {
    const onAnswer = vi.fn();
    render({
      pending: {
        task: '/rewrite',
        questions: ['What should I optimize for?'],
      },
      onAnswer,
    });

    const input = container.querySelector<HTMLInputElement>('.clarification-custom input');
    const submit = container.querySelector<HTMLButtonElement>('.clarification-custom button');
    act(() => {
      setInputValue(input!, 'Make it executive-ready');
      input!.dispatchEvent(new Event('input', { bubbles: true }));
    });
    act(() => submit?.click());
    expect(onAnswer).toHaveBeenCalledWith('Make it executive-ready');
  });
});

function setInputValue(input: HTMLInputElement, value: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
  setter?.call(input, value);
}
