// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { quickActionsForSurface, type QuickAction, type Surface } from '@ge/contracts';
import { QuickActionBar, type QuickActionBarProps } from './QuickActionBar.js';

/**
 * Behavioral tests for the quick-action chip row. The load-bearing logic: it renders exactly the
 * surface's `quickActionsForSurface(...)` catalog (capability closure, ADR-0006), each chip carries
 * its `output`/`intent` so the parent can route `chat` vs the write/annotation gate, clicking a chip
 * hands the whole action back, and the busy flag disables every chip.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

function render(props: Partial<QuickActionBarProps> = {}): { onAction: ReturnType<typeof vi.fn> } {
  const onAction = vi.fn();
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(
      createElement(QuickActionBar, {
        surface: 'word' as Surface,
        busy: false,
        onAction,
        ...props,
      }),
    );
  });
  return { onAction };
}

function chips(): HTMLButtonElement[] {
  return [...container.querySelectorAll<HTMLButtonElement>('button.quick-action')];
}

function chipLabels(): string[] {
  return chips().map((c) => c.querySelector('.quick-action-main')?.textContent ?? '');
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('QuickActionBar', () => {
  it('renders exactly the surface catalog, in order, by label', () => {
    render({ surface: 'word' });
    const expected = quickActionsForSurface('word').map((a) => a.label);
    expect(chipLabels()).toEqual(expected);
  });

  it('renders a different catalog per surface', () => {
    render({ surface: 'excel' });
    const expected = quickActionsForSurface('excel').map((a) => a.label);
    expect(chipLabels()).toEqual(expected);
    // Excel offers a range-summary action that Word does not.
    expect(chipLabels().some((label) => label === 'Summarize this range')).toBe(true);
    expect(expected).not.toEqual(quickActionsForSurface('word').map((a) => a.label));
  });

  it('narrows the catalog by allowed intents (capability closure)', () => {
    render({ surface: 'word', allowedIntents: ['ask', 'summarize', 'explain'] });
    const labels = chipLabels();
    const expected = quickActionsForSurface('word', ['ask', 'summarize', 'explain']).map(
      (a) => a.label,
    );
    expect(labels).toEqual(expected);
    // No annotation/write verb survives a chat-only closure.
    expect(chips().every((c) => c.getAttribute('data-output') === 'chat')).toBe(true);
  });

  it('tags each chip with its output and intent so the parent can route the gate', () => {
    render({ surface: 'word' });
    const byLabel = (label: string): HTMLButtonElement =>
      chips().find((c) => c.querySelector('.quick-action-main')?.textContent === label)!;
    const tighten = quickActionsForSurface('word').find((a) => a.id === 'tighten') as QuickAction;
    expect(tighten.intent).toBe('rewrite');
    const el = byLabel(tighten.label);
    expect(el.getAttribute('data-output')).toBe('write');
    expect(el.getAttribute('data-intent')).toBe('rewrite');
    expect(el.getAttribute('data-action-id')).toBe('tighten');
  });

  it('hands the full action back to onAction on click', () => {
    const { onAction } = render({ surface: 'word' });
    const first = quickActionsForSurface('word')[0];
    act(() => chips()[0]?.click());
    expect(onAction).toHaveBeenCalledTimes(1);
    expect(onAction).toHaveBeenCalledWith(first);
  });

  it('disables every chip while busy', () => {
    render({ surface: 'word', busy: true });
    expect(chips().length).toBeGreaterThan(0);
    expect(chips().every((c) => c.disabled)).toBe(true);
  });

  it('hides actions promoted into the contextual command center', () => {
    render({ surface: 'excel', excludeIds: ['create-chart', 'summarize-range'] });
    expect(chipLabels()).not.toContain('Create a chart');
    expect(chipLabels()).not.toContain('Summarize this range');
    expect(chipLabels()).toContain('Find anomalies / outliers');
  });
});
