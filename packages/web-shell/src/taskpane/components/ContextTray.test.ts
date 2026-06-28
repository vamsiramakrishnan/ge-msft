// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { ContextTray } from './ContextTray.js';
import type { ContextChip } from '../../controller.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement | undefined;
let root: Root | undefined;

afterEach(() => {
  if (root) act(() => root!.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

function render(chips: ContextChip[], onToggle = vi.fn(), onReveal = vi.fn()): HTMLDivElement {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root!.render(
      createElement(ContextTray, {
        chips,
        onToggle,
        onReveal,
        onRefresh: vi.fn(),
      }),
    );
  });
  return container;
}

describe('ContextTray', () => {
  it('renders revealable host context with a separate open control', () => {
    const onToggle = vi.fn();
    const onReveal = vi.fn();
    const chip: ContextChip = {
      id: 'xl:Sales!A1:C4',
      title: 'Sales!A1:C4',
      kind: 'range',
      attached: false,
      revealable: true,
    };
    const el = render([chip], onToggle, onReveal);

    const open = el.querySelector<HTMLButtonElement>('.open-host');
    expect(open).not.toBeNull();
    act(() => open?.click());

    expect(onReveal).toHaveBeenCalledWith('xl:Sales!A1:C4');
    expect(onToggle).not.toHaveBeenCalled();
  });

  it('does not render an open control for non-revealable context', () => {
    const chip: ContextChip = {
      id: 'doc:1',
      title: 'Document',
      kind: 'document',
      attached: false,
    };
    const el = render([chip]);

    expect(el.querySelector('.open-host')).toBeNull();
  });
});
