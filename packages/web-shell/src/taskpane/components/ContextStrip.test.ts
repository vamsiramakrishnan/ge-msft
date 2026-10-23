// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { ContextStrip } from './ContextStrip.js';
import type { ContextChip } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

function mount(chips: ContextChip[]): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(ContextStrip, {
        chips,
        disabled: false,
        onToggle: vi.fn(),
        onReveal: vi.fn(),
        onRefresh: vi.fn(),
      }),
    );
  });
}

const chips: ContextChip[] = [
  { id: 'a', title: 'Notebook: Vendor Risk', kind: 'document', attached: true },
  { id: 'b', title: 'SharePoint · Contracts', kind: 'document', attached: true },
  { id: 'c', title: 'Selection', kind: 'selection', attached: false },
];

describe('ContextStrip', () => {
  it('starts closed: one line naming the first source, no chips', () => {
    mount(chips);
    const toggle = container.querySelector<HTMLButtonElement>('.context-strip-toggle');
    expect(toggle?.textContent).toContain('Notebook: Vendor Risk');
    expect(toggle?.textContent).toContain('+1');
    expect(toggle?.textContent).not.toContain('available');
    expect(toggle?.getAttribute('aria-expanded')).toBe('false');
    expect(container.querySelectorAll('.smart-chip').length).toBe(0);
  });

  it('shows a context notice even while the dropdown is closed', () => {
    mount([
      ...chips,
      {
        id: 'xl:Sales!L:L',
        title: 'Sales!L:L',
        kind: 'range',
        attached: false,
        notice: 'Your selection is too large to read.',
      },
    ]);
    expect(container.querySelector('.context-notice')?.textContent).toContain(
      'Your selection is too large to read.',
    );
  });

  it('opens like a dropdown to show the attached chips', () => {
    mount(chips);
    act(() => container.querySelector<HTMLButtonElement>('.context-strip-toggle')?.click());
    expect(container.textContent).toContain('Notebook: Vendor Risk');
    expect(container.querySelectorAll('.smart-chip').length).toBe(2);
  });

  it('lists every source from "Add sources" inside the dropdown', () => {
    mount(chips);
    act(() => container.querySelector<HTMLButtonElement>('.context-strip-toggle')?.click());
    const add = [...container.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === 'Add sources (1)',
    );
    act(() => add?.click());
    expect(container.querySelectorAll('.smart-chip').length).toBe(3);
  });

  it('adds no hint line when nothing is attached, so the pane keeps its space', () => {
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
    act(() => {
      root.render(
        createElement(ContextStrip, {
          chips: [],
          disabled: false,
          onToggle: vi.fn(),
          onReveal: vi.fn(),
          onRefresh: vi.fn(),
        }),
      );
    });
    expect(container.textContent).not.toContain('Add sources to focus the answer.');
    expect(container.querySelector('.context-empty')).toBeNull();
  });
});
