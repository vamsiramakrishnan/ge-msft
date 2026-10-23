// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import { PanelController } from '../../controller.js';
import { installFakeExcel, type ExcelSimulator } from '../../test-harness/index.js';
import { selectBridge } from '../select-bridge.js';
import { App } from './App.js';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
let sim: ExcelSimulator;
let root: Root;
let container: HTMLDivElement;
let controller: PanelController;
let session: AssistSession;

beforeEach(() => {
  vi.spyOn(window, 'focus').mockImplementation(() => {});
});
afterEach(() => {
  controller?.cancel();
  act(() => root?.unmount());
  session?.dispose();
  container?.remove();
  sim?.restore();
  vi.restoreAllMocks();
});

function mount(initialModelId?: string): Record<string, unknown>[] {
  sim = installFakeExcel();
  const bodies: Record<string, unknown>[] = [];
  const client = new StreamAssistClient(
    { getAccessToken: async () => 'test-token' },
    { assistant: { project: 'test', location: 'global', engine: 'test' }, modelId: initialModelId },
    async (_url, init) => {
      bodies.push(JSON.parse(String(init?.body)) as Record<string, unknown>);
      return new Response(
        JSON.stringify([
          {
            answer: {
              state: 'SUCCEEDED',
              replies: [{ groundedContent: { content: { text: 'Ready.' } } }],
            },
          },
        ]),
        { status: 200 },
      );
    },
  );
  const bridge = selectBridge('excel')!;
  session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
  });
  controller = new PanelController(session, bridge);
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() =>
    root.render(
      createElement(App, {
        controller,
        surface: 'excel',
        initialModelId,
        onModelChange: (id) => client.setModelId(id),
      }),
    ),
  );
  act(() => container.querySelector<HTMLButtonElement>('[data-panel="settings"]')!.click());
  return bodies;
}

function select(): HTMLSelectElement {
  return container.querySelector<HTMLSelectElement>('select[aria-label="Model"]')!;
}

describe('Settings model selection', () => {
  it('starts with the engine default and sends the selected model on the next request', async () => {
    const bodies = mount();
    expect(select()).not.toBeNull();
    expect(select().value).toBe('');
    expect([...select().options].map((o) => o.textContent)).toEqual([
      'Default engine',
      'Gemini 4 Argon',
    ]);
    for (const modelId of ['gemini-4-argon', '']) {
      act(() => {
        select().value = modelId;
        select().dispatchEvent(new Event('change', { bubbles: true }));
      });
      await act(async () => {
        await controller.send('What is in this range?');
      });
      expect(bodies.at(-1)?.generationSpec).toEqual(
        modelId ? { modelId: 'gemini-4-argon' } : undefined,
      );
    }
  });

  it('shows the configured Argon model and locks selection while a write awaits approval', async () => {
    mount('gemini-4-argon');
    expect(select()).not.toBeNull();
    expect(select().value).toBe('gemini-4-argon');
    let running: Promise<unknown>;
    await act(async () => {
      running = controller.runDirectCommands('set A1 7');
    });
    expect(select().disabled).toBe(true);
    const deadline = Date.now() + 1000;
    while (
      !controller.getState().pendingPlan &&
      !controller.getState().pendingWrite &&
      Date.now() < deadline
    ) {
      await act(async () => {
        await new Promise((resolve) => setTimeout(resolve, 5));
      });
    }
    expect(controller.getState().pendingPlan ?? controller.getState().pendingWrite).toBeDefined();
    expect(select().disabled).toBe(true);
    await act(async () => {
      controller.cancel();
      await running;
    });
    expect(select().disabled).toBe(false);
  });
});
