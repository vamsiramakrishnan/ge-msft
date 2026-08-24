// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { quickActionsForSurface } from '@ge/contracts';
import {
  SurfaceCommandCenter,
  surfacePrimaryActions,
  type SurfaceCommandCenterProps,
} from './SurfaceCommandCenter.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement | undefined;
let root: Root | undefined;

function render(surface: SurfaceCommandCenterProps['surface'] = 'excel'): {
  onAction: ReturnType<typeof vi.fn>;
} {
  const onAction = vi.fn();
  container = document.createElement('div');
  document.body.appendChild(container);
  const nextRoot = createRoot(container);
  root = nextRoot;
  act(() => {
    nextRoot.render(
      createElement(SurfaceCommandCenter, {
        surface,
        busy: false,
        hasGate: false,
        attachedCount: 2,
        availableCount: 3,
        messageCount: 4,
        proposalCount: 1,
        onAction,
      }),
    );
  });
  return { onAction };
}

function buttons(): HTMLButtonElement[] {
  return [...(container?.querySelectorAll<HTMLButtonElement>('button.surface-action') ?? [])];
}

afterEach(() => {
  const existingRoot = root;
  if (existingRoot) act(() => existingRoot.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

describe('SurfaceCommandCenter', () => {
  it('promotes the best Excel actions into the primary area', () => {
    expect(surfacePrimaryActions('excel').map((action) => action.id)).toEqual([
      'create-chart',
      'summarize-range',
      'find-anomalies',
    ]);
  });

  it('promotes distinct primary workflows for Word, PowerPoint, and Outlook', () => {
    expect(surfacePrimaryActions('word').map((action) => action.id)).toEqual([
      'tighten',
      'comment-on-issues',
      'review-against',
    ]);
    expect(surfacePrimaryActions('powerpoint').map((action) => action.id)).toEqual([
      'draft-slide',
      'draft-section',
      'redesign',
    ]);
    expect(surfacePrimaryActions('outlook').map((action) => action.id)).toEqual([
      'recent-mail-briefing',
      'prepare-next-meeting',
      'draft-reply',
    ]);
    expect(surfacePrimaryActions('onenote').map((action) => action.id)).toEqual([
      'synthesize-page',
      'discover-sources',
      'audio-overview',
    ]);
    expect(surfacePrimaryActions('teams').map((action) => action.id)).toEqual([
      'live-notes',
      'action-items',
      'catch-up-teams',
    ]);
  });

  it('honors capability closure when choosing primary actions', () => {
    expect(
      surfacePrimaryActions('excel', ['ask', 'summarize', 'explain']).map((a) => a.id),
    ).toEqual(['summarize-range', 'explain-formula', 'summarize-this']);
  });

  it('renders host-specific copy and compact state metrics', () => {
    render('outlook');
    expect(container?.querySelector('.surface-title')?.textContent).toBe('Outlook workspace');
    expect(container?.querySelector('.surface-details-popover')?.textContent).toContain('Threads');
    expect(container?.querySelector('.surface-metrics')?.textContent).toContain('2 attached');
    expect(container?.querySelector('.surface-metrics')?.textContent).toContain('1 proposals');
  });

  it('hands the full promoted action back on click', () => {
    const { onAction } = render('excel');
    act(() => buttons()[0]?.click());
    expect(onAction).toHaveBeenCalledWith(
      quickActionsForSurface('excel').find((action) => action.id === 'create-chart'),
    );
  });

  it('disables promoted actions while a gate is active', () => {
    const onAction = vi.fn();
    container = document.createElement('div');
    document.body.appendChild(container);
    const nextRoot = createRoot(container);
    root = nextRoot;
    act(() => {
      nextRoot.render(
        createElement(SurfaceCommandCenter, {
          surface: 'word',
          busy: false,
          hasGate: true,
          attachedCount: 0,
          availableCount: 0,
          messageCount: 0,
          proposalCount: 0,
          onAction,
        }),
      );
    });
    expect(buttons().length).toBe(3);
    expect(buttons().every((button) => button.disabled)).toBe(true);
    expect(container?.querySelector('.surface-state')?.textContent).toBe('Gate');
  });
});
