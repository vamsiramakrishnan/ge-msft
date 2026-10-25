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
  it('hides the exact command until the user opens the effect', () => {
    render({ effects: [effect], summary: '1 write' });
    const head = container.querySelector<HTMLButtonElement>('.plan-effect-head');
    expect(head?.getAttribute('aria-expanded')).toBe('false');
    expect(container.textContent).not.toContain('set Sales!F2 =C2-D2');
    expect(head?.textContent).toContain('Sales!F2');

    act(() => head?.click());
    expect(head?.getAttribute('aria-expanded')).toBe('true');
    expect(container.querySelector('.plan-effect-detail .cmd')?.textContent).toBe(
      'set Sales!F2 =C2-D2',
    );
  });

  it('previews what each change writes while collapsed, without the command', () => {
    render({ effects: [effect], summary: '1 write' });
    const head = container.querySelector<HTMLButtonElement>('.plan-effect-head');
    expect(head?.querySelector('.effect-change')?.textContent).toContain('=C2-D2');
    expect(container.textContent).not.toContain('set Sales!F2 =C2-D2');
  });

  it('names each new slide by its title, and shows its title and bullets when opened', () => {
    const slide = (id: string, title: string, bullets: string[]): PlanEffect => ({
      command: '',
      approvalClass: approvalClassOf('insert-slide'),
      reversible: isReversibleKind('insert-slide'),
      request: {
        changeId: asChangeId(id),
        kind: 'insert-slide',
        surface: 'powerpoint',
        params: { slide: { title, bullets } },
      },
    });
    render({
      effects: [
        slide('s1', 'Latar belakang', ['Adopsi AI meningkat']),
        slide('s2', 'Risiko utama', ['Bias model', 'Kebocoran data']),
      ],
      summary: '2 insert-slides',
    });
    const heads = [...container.querySelectorAll<HTMLButtonElement>('.plan-effect-head')];
    expect(heads.map((head) => head.querySelector('.effect-change')?.textContent)).toEqual([
      'Latar belakang',
      'Risiko utama',
    ]);
    act(() => heads[1]?.click());
    expect(container.querySelector('.plan-effect-detail .cmd')?.textContent).toBe(
      'slide "Risiko utama" "Bias model" "Kebocoran data"',
    );
  });

  it('shows recipients and a Review badge on a high-impact change', () => {
    const recipients: PlanEffect = {
      command: '/set-recipients bcc=attacker@example.com',
      approvalClass: approvalClassOf('set-recipients'),
      reversible: isReversibleKind('set-recipients'),
      request: {
        changeId: asChangeId('c2'),
        kind: 'set-recipients',
        surface: 'outlook',
        params: { mail: { bcc: ['attacker@example.com'] } },
      },
    };
    render({ effects: [recipients], summary: '1 change' });
    const head = container.querySelector<HTMLButtonElement>('.plan-effect-head');
    expect(head?.textContent).toContain('bcc: attacker@example.com');
    expect(head?.querySelector('.effect-review')?.textContent).toBe('Check');
  });

  it('shows recipients even when the dry-run has its own preview', () => {
    const recipients: PlanEffect = {
      command: '/set-recipients bcc=attacker@example.com',
      approvalClass: approvalClassOf('set-recipients'),
      reversible: isReversibleKind('set-recipients'),
      request: {
        changeId: asChangeId('c3'),
        kind: 'set-recipients',
        surface: 'outlook',
        params: { mail: { bcc: ['attacker@example.com'] } },
      },
      dryRun: { after: 'Recipients updated' },
    };
    render({ effects: [recipients], summary: '1 change' });
    expect(container.querySelector('.effect-change')?.textContent).toContain(
      'bcc: attacker@example.com',
    );
  });

  it('reveals effect targets through a navigation-only callback', () => {
    const onRevealTarget = vi.fn();
    render({ effects: [effect], summary: '1 write' }, onRevealTarget);

    act(() => container.querySelector<HTMLButtonElement>('.plan-effect-head')?.click());
    const target = container.querySelector<HTMLButtonElement>('.host-target-link');
    expect(target?.textContent).toBe('Sales!F2');

    act(() => target?.click());
    expect(onRevealTarget).toHaveBeenCalledWith('Sales!F2');
  });

  it('says what changes in plain words, in the conversation language', () => {
    render({ effects: [effect], summary: '1 write' });
    expect(container.textContent).toContain('Review this change');
    expect(container.querySelector('.plan-effect-kind')?.textContent).toBe('Cells');
    expect(container.textContent).not.toContain('write-cells');
    expect(container.textContent).not.toContain('1 write');
    expect([...container.querySelectorAll('.act button')].map((b) => b.textContent)).toEqual([
      'Apply',
      'Cancel',
    ]);
  });
});
