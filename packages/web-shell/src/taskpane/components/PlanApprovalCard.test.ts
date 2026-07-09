// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { asChangeId, approvalClassOf, isReversibleKind } from '@ge/contracts';
import { PlanApprovalCard } from './PlanApprovalCard.js';
import type { PendingPlan, PlanEffect } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

const effect: PlanEffect = {
  command: 'set Sales!F2 =C2-D2',
  approvalClass: approvalClassOf('write-cells'),
  reversible: isReversibleKind('write-cells'),
  request: {
    changeId: asChangeId('c1'),
    kind: 'write-cells',
    surface: 'excel',
    params: { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
  },
  dryRun: { target: 'Sales!F2', before: '', after: '=C2-D2' },
};

function render(plan: PendingPlan, onRevealTarget = vi.fn()): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(PlanApprovalCard, {
        plan,
        onRevealTarget,
        onApprove: vi.fn(),
        onReject: vi.fn(),
      }),
    );
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('PlanApprovalCard', () => {
  it('reveals effect targets through a navigation-only callback', () => {
    const onRevealTarget = vi.fn();
    render({ effects: [effect], summary: '1 write' }, onRevealTarget);

    act(() => container.querySelector<HTMLButtonElement>('.plan-effect-head')?.click());
    const target = container.querySelector<HTMLButtonElement>('.host-target-link');
    expect(target?.textContent).toBe('Sales!F2');

    act(() => target?.click());
    expect(onRevealTarget).toHaveBeenCalledWith('Sales!F2');
  });
});
