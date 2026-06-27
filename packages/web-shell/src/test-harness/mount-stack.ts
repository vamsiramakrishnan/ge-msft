/**
 * The full-stack mount helper for integration tests: install a fake host → pick the REAL bridge via
 * `selectBridge` → wire a REAL {@link AssistSession} (over a scripted fake client) → REAL
 * {@link PanelController} → render the REAL `<App/>` (the `createRoot`+`act` pattern from
 * `app-render.test.ts`). Nothing here is a mock except the host object model and the model stream,
 * so an assertion against the rendered DOM is an assertion about the whole client stack.
 *
 * jsdom-only (it renders React). Import from a `// @vitest-environment jsdom` test file.
 */

import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import type { Surface, UnitDescriptor } from '@ge/contracts';
import { AssistSession } from '@ge/runtime';
import type { DocBridge } from '@ge/runtime';
import type { TriggerRegistry } from '@ge/triggers';
import type { StreamAssistClient } from '@ge/gemini-client';
import { selectBridge } from '../taskpane/select-bridge.js';
import { PanelController } from '../controller.js';
import { App } from '../taskpane/components/App.js';

/** Options for {@link mountStack}. */
export interface MountStackOptions {
  surface: Surface;
  /** The model stream: a `StreamAssistClient` or the `{ client }` wrapper `scriptedClient` returns. */
  client: StreamAssistClient | { client: StreamAssistClient };
  /** Override the research-unit descriptor (defaults to a bare unit for the surface). */
  unit?: UnitDescriptor;
  /** Optional actuation gate / event registry wired into the session. */
  triggers?: TriggerRegistry;
}

/** Accept either a raw client or the `{ client, queries }` wrapper from `scriptedClient`. */
function resolveClient(client: MountStackOptions['client']): StreamAssistClient {
  return 'client' in client ? client.client : client;
}

/** The mounted stack: the live objects plus the rendered container and a teardown. */
export interface MountedStack {
  bridge: DocBridge;
  session: AssistSession;
  controller: PanelController;
  container: HTMLDivElement;
  root: Root;
  /** Flush React state updates inside `act` (await microtasks the stack scheduled). */
  flush(): Promise<void>;
  /**
   * Pump the event loop inside `act` until `predicate(controller.getState())` holds (or `timeoutMs`
   * elapses). The command loop suspends on async stream consumption between gates, so an integration
   * test must wait for a gate (e.g. `pendingPlan`) to be staged rather than assume a fixed number of
   * microtasks. Rejects if the predicate never holds within the budget.
   */
  waitFor(
    predicate: (state: ReturnType<PanelController['getState']>) => boolean,
    timeoutMs?: number,
  ): Promise<void>;
  /**
   * Run a controller-mutating action (e.g. `() => controller.approvePlan()`) inside `act`, flushing
   * the synchronous + immediate-microtask React updates it triggers — so the rendered DOM stays
   * warning-free. Use for the discrete user gestures the UI exposes (approve/reject/send/run-kickoff).
   */
  act(fn: () => void): Promise<void>;
  /** Unmount React + remove the container. Does NOT restore host globals (the simulator owns that). */
  unmount(): void;
}

/** A bare unit descriptor for `surface` — no notebook, no connectors, just the surface context. */
function bareUnit(surface: Surface): UnitDescriptor {
  return { connectors: [], surfaceContext: { kind: surface } };
}

/**
 * Mount the real stack over a fake host (already installed by the caller via `installFake*`).
 * Renders `<App/>` and returns the live controller/session/bridge so a test can drive the controller
 * (e.g. `runCommands`, `approvePlan`) and assert both the rendered DOM and the mutated fake host.
 */
export function mountStack(opts: MountStackOptions): MountedStack {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

  const bridge = selectBridge(opts.surface);
  if (!bridge) throw new Error(`mountStack: no bridge for surface "${opts.surface}"`);

  const session = new AssistSession(bridge, resolveClient(opts.client), {
    unit: opts.unit ?? bareUnit(opts.surface),
    ...(opts.triggers ? { triggers: opts.triggers } : {}),
  });
  const controller = new PanelController(session, bridge);

  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(createElement(App, { controller, surface: opts.surface }));
  });

  const flush = async (): Promise<void> => {
    await act(async () => {
      // Let scheduled microtasks (the command loop's awaits + controller state updates) settle
      // inside act(). Several drains because the loop chains multiple awaits before it gates.
      for (let i = 0; i < 20; i++) await Promise.resolve();
    });
  };

  const waitFor = async (
    predicate: (state: ReturnType<PanelController['getState']>) => boolean,
    timeoutMs = 1000,
  ): Promise<void> => {
    const deadline = Date.now() + timeoutMs;
    while (!predicate(controller.getState())) {
      if (Date.now() > deadline) {
        throw new Error('mountStack.waitFor: predicate not satisfied within timeout');
      }
      // Advance the suspended async command loop one macrotask, flushing React updates each tick so
      // the rendered DOM tracks state. Kept OUTSIDE a long-lived act() so the loop actually runs.
      await act(async () => {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
      });
    }
  };

  const actGesture = async (fn: () => void): Promise<void> => {
    await act(async () => {
      fn();
      // Drain the immediate microtasks the gesture scheduled (e.g. the resolved approval promise
      // letting the loop take its next step) so their React updates land inside this act().
      for (let i = 0; i < 5; i++) await Promise.resolve();
    });
  };

  const unmount = (): void => {
    act(() => root.unmount());
    container.remove();
  };

  return {
    bridge,
    session,
    controller,
    container,
    root,
    flush,
    waitFor,
    act: actGesture,
    unmount,
  };
}
