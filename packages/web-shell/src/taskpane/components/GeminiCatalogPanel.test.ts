// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { DiscoveryCatalogClient, GeminiCatalog } from '@ge/gemini-client';
import { GeminiCatalogPanel } from './GeminiCatalogPanel.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const catalog: GeminiCatalog = {
  skills: [
    {
      name: 'projects/p/locations/global/agents/17573173582293271726',
      id: '17573173582293271726',
      label: 'm365-command-planner',
      mention: { label: 'm365-command-planner', uri: '17573173582293271726' },
      suggestedRoute: 'planner',
    },
    {
      name: 'projects/p/locations/global/agents/7404511736383961129',
      id: '7404511736383961129',
      label: 'm365-surface-commander',
      mention: { label: 'm365-surface-commander', uri: '7404511736383961129' },
      suggestedRoute: 'command',
    },
  ],
  dataStores: [
    {
      name: 'projects/p/locations/global/collections/default_collection/dataStores/msft-onedrive',
      id: 'msft-onedrive',
      label: 'OneDrive files',
      suggested: true,
    },
  ],
  connectors: [],
};

let container: HTMLDivElement | undefined;
let root: Root | undefined;

function clientWith(listCatalog: DiscoveryCatalogClient['listCatalog']): DiscoveryCatalogClient {
  return { listCatalog } as DiscoveryCatalogClient;
}

function render(props: {
  catalogClient: DiscoveryCatalogClient;
  disabled?: boolean;
  onApply?: ReturnType<typeof vi.fn>;
}): ReturnType<typeof vi.fn> {
  const onApply = props.onApply ?? vi.fn();
  container = document.createElement('div');
  document.body.appendChild(container);
  const nextRoot = createRoot(container);
  root = nextRoot;
  act(() => {
    nextRoot.render(
      createElement(GeminiCatalogPanel, {
        catalogClient: props.catalogClient,
        disabled: props.disabled,
        onApply,
      }),
    );
  });
  return onApply;
}

afterEach(() => {
  if (root) act(() => root!.unmount());
  container?.remove();
  root = undefined;
  container = undefined;
});

describe('GeminiCatalogPanel', () => {
  it('does not load or apply routing while disabled', () => {
    const listCatalogMock = vi.fn();
    const onApply = render({
      catalogClient: clientWith(
        listCatalogMock as unknown as DiscoveryCatalogClient['listCatalog'],
      ),
      disabled: true,
    });

    expect(listCatalogMock).not.toHaveBeenCalled();
    expect(onApply).not.toHaveBeenCalled();
    expect(container?.querySelector<HTMLButtonElement>('.mini-btn')?.disabled).toBe(true);
  });

  it('does not apply defaults if the panel becomes disabled before catalog load resolves', async () => {
    let resolveCatalog: (value: GeminiCatalog) => void = () => {};
    const listCatalogMock = vi.fn(
      () =>
        new Promise<GeminiCatalog>((resolve) => {
          resolveCatalog = resolve;
        }),
    );
    const client = clientWith(listCatalogMock as unknown as DiscoveryCatalogClient['listCatalog']);
    const onApply = vi.fn();
    render({ catalogClient: client, onApply });

    expect(listCatalogMock).toHaveBeenCalledTimes(1);
    await act(async () => {
      root?.render(
        createElement(GeminiCatalogPanel, { catalogClient: client, disabled: true, onApply }),
      );
      resolveCatalog(catalog);
      await Promise.resolve();
    });

    expect(onApply).not.toHaveBeenCalled();
  });

  it('applies discovered default planner, command skill, and connectors when enabled', async () => {
    const listCatalogMock = vi.fn(async () => catalog);
    const onApply = render({
      catalogClient: clientWith(
        listCatalogMock as unknown as DiscoveryCatalogClient['listCatalog'],
      ),
      onApply: vi.fn(),
    });

    await act(async () => {
      await Promise.resolve();
    });

    expect(onApply).toHaveBeenCalledWith(
      expect.objectContaining({
        plannerSkill: expect.objectContaining({ label: 'm365-command-planner' }),
        commandSkill: expect.objectContaining({ label: 'm365-surface-commander' }),
        dataStores: [expect.objectContaining({ label: 'OneDrive files' })],
      }),
    );
    expect(container?.textContent).toContain('2 skills');
    // With no connectors listed, the summary keeps its original wording.
    expect(container?.textContent).toContain('1 connectors');
  });

  it('renders the connector board (lamp plus state word) above the store checkboxes', async () => {
    const withConnectors: GeminiCatalog = {
      ...catalog,
      dataStores: [
        ...catalog.dataStores,
        {
          name: 'projects/p/locations/global/collections/sharepoint-connector/dataStores/sp-files',
          id: 'sp-files',
          label: 'SharePoint files',
        },
      ],
      connectors: [
        {
          name: 'projects/p/locations/global/collections/sharepoint-connector',
          id: 'sharepoint-connector',
          label: 'SharePoint',
          source: 'sharepoint',
          state: 'failed',
          lastSyncTime: '2026-07-01T10:00:00Z',
        },
      ],
    };
    const listCatalogMock = vi.fn(async () => withConnectors);
    render({
      catalogClient: clientWith(
        listCatalogMock as unknown as DiscoveryCatalogClient['listCatalog'],
      ),
    });
    await act(async () => {
      await Promise.resolve();
    });

    // Summary switches to the three-part wording when connectors exist.
    expect(container?.textContent).toContain('2 skills · 1 connectors · 2 stores');

    // Expand the settings grid.
    const buttons = [...(container?.querySelectorAll<HTMLButtonElement>('.mini-btn') ?? [])];
    const edit = buttons.find((button) => button.textContent === 'Edit');
    expect(edit).toBeDefined();
    act(() => edit!.click());

    const board = container?.querySelector('.connector-board');
    expect(board).toBeTruthy();
    const row = board?.querySelector('.connector-row');
    expect(row?.querySelector('.connector-lamp')?.getAttribute('data-tone')).toBe('err');
    expect(row?.querySelector('.connector-state')?.textContent).toBe('failed');
    expect(row?.querySelector('.connector-label')?.textContent).toBe('SharePoint');
    expect(row?.querySelector('.connector-source')?.textContent).toBe('sharepoint');
    expect(row?.querySelector('.connector-stores')?.textContent).toBe('1 store');
    expect(row?.querySelector('.connector-sync')?.textContent).toBeTruthy();

    // The board sits ABOVE the existing checkbox fieldset, which stays intact.
    const grid = container?.querySelector('.catalog-grid');
    const children = [...(grid?.children ?? [])];
    expect(children.findIndex((el) => el.classList.contains('connector-board'))).toBeLessThan(
      children.findIndex((el) => el.classList.contains('catalog-connectors')),
    );
    expect(grid?.querySelectorAll('.catalog-check input[type="checkbox"]')).toHaveLength(2);
  });
});
