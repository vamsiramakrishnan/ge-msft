// @vitest-environment jsdom
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

/**
 * Behavioral test for the preview harness bootstrap + the `Preview` toolbar component. The module
 * self-mounts the real <App/> over the fake controller into `#root` on import, so by providing a
 * `#root` element and importing fresh, we drive: the default everything-on snapshot, the surface
 * switcher (re-rendering App with a new `data-surface`), the per-card toggles (each removing its
 * card from the derived `PanelState`), and the "Idle / empty" preset. No Office host, no network.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: HTMLDivElement;

beforeEach(() => {
  vi.resetModules();
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  root = document.createElement('div');
  root.id = 'root';
  document.body.appendChild(root);
});

afterEach(() => {
  vi.restoreAllMocks();
  document.body.innerHTML = '';
});

/** Import the module fresh so its top-level `createRoot(#root).render(<Preview/>)` runs. */
async function mountPreview(): Promise<void> {
  const { act } = await import('react');
  await act(async () => {
    await import('./preview.js');
  });
}

function click(el: Element | null | undefined): Promise<void> {
  return import('react').then(({ act }) =>
    act(async () => {
      (el as HTMLElement | null)?.click();
    }),
  );
}

describe('preview bootstrap + Preview toolbar', () => {
  it('mounts the real panel into #root with every card on by default', async () => {
    await mountPreview();
    // The default surface is Word.
    const panel = root.querySelector('.panel');
    expect(panel?.getAttribute('data-surface')).toBe('word');
    // Everything-on: thread, context tray, suggestions, run steps, plan, write, proposals, error.
    expect(root.querySelector('.thread')).not.toBeNull();
    expect(root.querySelector('.plan-approval')).not.toBeNull();
    expect(root.querySelector('.approval')).not.toBeNull();
    expect(root.querySelector('.proposals')).not.toBeNull();
    expect(root.querySelector('.panel-error')).not.toBeNull();
    expect(root.querySelector('.skills')).not.toBeNull();
  });

  it('switches the host surface via the toolbar, re-rendering the panel', async () => {
    await mountPreview();
    const excelChip = Array.from(root.querySelectorAll<HTMLButtonElement>('.preview-chip')).find(
      (b) => b.textContent === 'excel',
    );
    expect(excelChip).toBeTruthy();
    await click(excelChip);
    expect(root.querySelector('.panel')?.getAttribute('data-surface')).toBe('excel');
    expect(excelChip?.getAttribute('aria-pressed')).toBe('true');
  });

  it('removes a card from the panel when its toolbar toggle is switched off', async () => {
    await mountPreview();
    // The plan-approval card is present in the default everything-on state.
    expect(root.querySelector('.plan-approval')).not.toBeNull();
    const planToggle = Array.from(root.querySelectorAll('.preview-toggle')).find((l) =>
      l.textContent?.includes('Plan'),
    );
    const checkbox = planToggle?.querySelector<HTMLInputElement>('input');
    await click(checkbox);
    // Toggling Plan off drops `pendingPlan` from the derived state, so the card unmounts.
    expect(root.querySelector('.plan-approval')).toBeNull();
  });

  it('collapses to the idle/empty preset, leaving only the thread', async () => {
    await mountPreview();
    const idle = Array.from(root.querySelectorAll<HTMLButtonElement>('.preview-btn')).find(
      (b) => b.textContent === 'Idle / empty',
    );
    await click(idle);
    // Idle preset keeps messages on but turns every other card off.
    expect(root.querySelector('.thread')).not.toBeNull();
    expect(root.querySelector('.plan-approval')).toBeNull();
    expect(root.querySelector('.approval')).toBeNull();
    expect(root.querySelector('.proposals')).toBeNull();
    expect(root.querySelector('.panel-error')).toBeNull();
    expect(root.querySelector('.skills')).toBeNull();
  });

  it('restores everything with the "All on" preset after idling', async () => {
    await mountPreview();
    const findBtn = (label: string): HTMLButtonElement | undefined =>
      Array.from(root.querySelectorAll<HTMLButtonElement>('.preview-btn')).find(
        (b) => b.textContent === label,
      );
    await click(findBtn('Idle / empty'));
    expect(root.querySelector('.plan-approval')).toBeNull();
    await click(findBtn('All on'));
    expect(root.querySelector('.plan-approval')).not.toBeNull();
    expect(root.querySelector('.proposals')).not.toBeNull();
  });
});
