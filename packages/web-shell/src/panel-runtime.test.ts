import { describe, expect, it, vi } from 'vitest';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, type DocBridge, type RuntimeExtension } from '@ge/runtime';
import type { ContextRef } from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { createMessageSendHandler } from '@ge/bridge-outlook';
import { connectPanelRuntime } from './panel-runtime.js';
import { APPLICATION_EXTENSIONS, createApplicationRuntime } from './runtime-extensions.js';
import { PanelController } from './controller.js';

function fixture(extensions: readonly RuntimeExtension[] = APPLICATION_EXTENSIONS) {
  const runtime = createApplicationRuntime(extensions);
  let emit: ((event: HostEvent) => void) | undefined;
  const unwatch = vi.fn();
  const list = vi.fn(async (): Promise<ContextRef[]> => []);
  const bridge: DocBridge = {
    surface: 'outlook',
    getCapabilities: () => ({ surface: 'outlook', contextKinds: [], actuations: [] }),
    listContext: list,
    resolveContext: async () => [],
    actuate: async (r) => ({ ok: true, changeId: r.changeId, kind: r.kind }),
    watch(listener) {
      emit = listener;
      return () => {
        emit = undefined;
        unwatch();
      };
    },
  };
  const fetch = vi.fn(async () => new Response('[]'));
  const client = new StreamAssistClient(
    { getAccessToken: async () => 'test' },
    { assistant: { project: 'p', location: 'global', engine: 'e' } },
    fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'outlook' } },
    hooks: runtime.hooks,
    triggers: runtime.triggers,
    primeOnHostEvent: false,
  });
  const controller = new PanelController(session, bridge);
  const connection = connectPanelRuntime({
    session,
    bridge,
    controller,
    triggers: runtime.triggers,
  });
  return {
    runtime,
    session,
    controller,
    connection,
    list,
    fetch,
    unwatch,
    emit: (event: HostEvent) => emit?.(event),
  };
}

describe('production event and extension wiring', () => {
  it('runs host message hooks, constructs context, and offers an action without a model call', async () => {
    const seen = vi.fn();
    const f = fixture([
      ...APPLICATION_EXTENSIONS,
      {
        id: 'tenant.mail',
        setup(api) {
          api.on({
            id: 'received',
            on: 'host:event',
            mode: 'observe',
            handle({ event }) {
              if (event.type === 'mail-received') seen(event.id);
            },
          });
        },
      },
    ]);
    f.connection.start();
    f.emit({ type: 'mail-received', id: 'message-1' });
    await f.connection.idle();
    expect(seen).toHaveBeenCalledWith('message-1');
    expect(f.controller.getState().suggestions[0]?.title).toBe('Catch up on this message');
    expect(f.session.model.hasPending).toBe(true);
    expect(f.list).toHaveBeenCalled();
    expect(f.fetch).not.toHaveBeenCalled();
    f.connection.dispose();
    f.runtime.dispose();
  });

  it('suppresses late suggestions after stop, restarts once, and removes every subscription', async () => {
    let finish!: () => void;
    let entered!: () => void;
    const started = new Promise<void>((r) => {
      entered = r;
    });
    const f = fixture([
      {
        id: 'slow',
        setup(api) {
          api.trigger({
            id: 'suggest',
            on: 'mail-received',
            handle: async () => {
              entered();
              await new Promise<void>((r) => {
                finish = r;
              });
              return { kind: 'suggest', title: 'Late' };
            },
          });
        },
      },
    ]);
    f.connection.start();
    f.emit({ type: 'mail-received', id: 'x' });
    await started;
    f.connection.stop();
    finish();
    await f.connection.idle();
    expect(f.controller.getState().suggestions).toEqual([]);
    expect(f.connection.receive({ type: 'mail-received', id: 'stopped' })).toBe(false);
    f.connection.start();
    f.connection.start();
    f.connection.dispose();
    f.connection.dispose();
    expect(f.unwatch).toHaveBeenCalledTimes(2);
    f.runtime.dispose();
  });

  it('turns opt-in automation outcomes into a visible invitation in the production pane', async () => {
    const f = fixture([
      {
        id: 'automation',
        setup(api) {
          api.trigger({
            id: 'mail',
            on: 'mail-received',
            handle: () => ({ kind: 'automate', query: 'Review this message' }),
          });
        },
      },
    ]);
    f.connection.start();
    expect(f.connection.receive({ type: 'mail-received', id: 'external-adapter-event' })).toBe(
      true,
    );
    await f.connection.idle();
    expect(f.controller.getState().suggestions[0]?.query).toBe('Review this message');
    expect(f.fetch).not.toHaveBeenCalled();
    f.connection.dispose();
    f.runtime.dispose();
  });

  it('installs the same required mail-send guards in a separate function-runtime registry', async () => {
    const extension: RuntimeExtension = {
      id: 'tenant.send',
      setup(api) {
        api.trigger({
          id: 'policy',
          on: 'mail-send',
          handle: () => ({ kind: 'block', reason: 'Review the external recipients.' }),
        });
      },
    };
    const pane = createApplicationRuntime([extension]);
    const command = createApplicationRuntime([extension]);
    const completed = vi.fn();
    await createMessageSendHandler(command.triggers)({ completed });
    expect(completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage: 'Review the external recipients.',
    });
    command.dispose();
    expect(pane.triggers.size).toBe(1);
    pane.dispose();
  });
});

it('defers a host context refresh that completes during a task, then refreshes after idle', async () => {
  const f = fixture();
  let finishListing!: (refs: ContextRef[]) => void;
  f.list.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        finishListing = resolve;
      }),
  );
  let releaseTask!: () => void;
  let entered!: () => void;
  const started = new Promise<void>((r) => {
    entered = r;
  });
  f.runtime.hooks.register({
    id: 'hold-task',
    on: 'message:received',
    mode: 'observe',
    handle: async () => {
      entered();
      await new Promise<void>((r) => {
        releaseTask = r;
      });
    },
  });
  f.connection.start();
  f.emit({ type: 'mail-received', id: 'first' });
  await f.connection.idle();
  const turn = f.controller.send('Question');
  await started;
  finishListing([{ id: 'stale', kind: 'selection', surface: 'outlook', title: 'Old context' }]);
  await Promise.resolve();
  await Promise.resolve();
  expect(f.controller.getState().chips).toEqual([]);
  releaseTask();
  await turn;
  expect(f.list).toHaveBeenCalledTimes(2);
  expect(f.controller.getState().chips).toEqual([]);
  f.connection.dispose();
  f.runtime.dispose();
});
