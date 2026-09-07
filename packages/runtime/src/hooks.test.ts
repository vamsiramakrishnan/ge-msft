import { afterEach, describe, expect, it, vi } from 'vitest';
import { RuntimeHooks, HookBlockedError, type RuntimeHook } from './hooks.js';
import { installRuntimeExtensions, type RuntimeExtensionApi } from './extensions.js';
import { TriggerRegistry } from '@ge/triggers';

const context = { taskId: 'task-1', surface: 'word' as const };
const payload = { mode: 'chat' as const, text: 'private customer content' };
afterEach(() => {
  vi.useRealTimers();
});

describe('RuntimeHooks', () => {
  it('orders hooks by priority then registration, and unregisters only the owned hook', async () => {
    const hooks = new RuntimeHooks();
    const calls: string[] = [];
    const add = (id: string, priority: number) =>
      hooks.register({
        id,
        on: 'message:received',
        mode: 'observe',
        priority,
        handle: () => {
          calls.push(id);
        },
      });
    add('last', -1);
    const off = add('first', 2);
    add('second', 2);
    await hooks.run('message:received', payload, context);
    expect(calls).toEqual(['first', 'second', 'last']);
    off();
    off();
    calls.length = 0;
    await hooks.run('message:received', payload, context);
    expect(calls).toEqual(['second', 'last']);
  });

  it('isolates payloads and observer errors, and never records prompt or exception content', async () => {
    const hooks = new RuntimeHooks();
    hooks.register({
      id: 'mutation',
      on: 'message:received',
      mode: 'observe',
      handle: (event) => {
        (event as { text: string }).text = 'changed';
      },
    });
    hooks.register({
      id: 'throws',
      on: 'message:received',
      mode: 'observe',
      handle: () => {
        throw new Error(payload.text);
      },
    });
    const seen = vi.fn();
    hooks.register({ id: 'next', on: 'message:received', mode: 'observe', handle: seen });
    hooks.subscribe(() => {
      throw new Error('broken diagnostics');
    });
    await hooks.run('message:received', payload, context);
    expect(seen.mock.calls[0]?.[0]).toEqual(payload);
    expect(hooks.records().map((r) => r.outcome)).toEqual(['error', 'error', 'continued']);
    expect(JSON.stringify(hooks.records())).not.toContain(payload.text);
    expect(payload.text).toBe('private customer content');
  });

  it('stops at a guard refusal and preserves the explicit reason', async () => {
    const hooks = new RuntimeHooks();
    hooks.register({
      id: 'policy',
      on: 'message:received',
      mode: 'guard',
      handle: () => ({ kind: 'block', reason: 'Select an approved source.' }),
    });
    const next = vi.fn();
    hooks.register({ id: 'next', on: 'message:received', mode: 'observe', handle: next });
    await expect(hooks.run('message:received', payload, context)).rejects.toThrow(
      'Select an approved source.',
    );
    expect(next).not.toHaveBeenCalled();
  });

  it('fails closed on guard timeout, cancels its signal, and rejects late context', async () => {
    vi.useFakeTimers();
    const hooks = new RuntimeHooks();
    let signal!: AbortSignal;
    let finish!: (value: { kind: 'continue' }) => void;
    hooks.register({
      id: 'hung',
      on: 'message:received',
      mode: 'guard',
      timeoutMs: 20,
      handle: (_, ctx) => {
        signal = ctx.signal;
        return new Promise((r) => {
          finish = r;
        });
      },
    });
    const pending = hooks.run('message:received', payload, context);
    const checked = expect(pending).rejects.toThrow(HookBlockedError);
    await vi.advanceTimersByTimeAsync(21);
    await checked;
    expect(signal.aborted).toBe(true);
    finish({ kind: 'continue' });
    await Promise.resolve();
    expect(hooks.records()).toHaveLength(1);
    expect(hooks.records()[0]?.outcome).toBe('timeout');
  });

  it('cancels during a handler and runs no subsequent handler', async () => {
    const hooks = new RuntimeHooks();
    const abort = new AbortController();
    let entered!: () => void;
    const started = new Promise<void>((r) => {
      entered = r;
    });
    hooks.register({
      id: 'waiting',
      on: 'message:received',
      mode: 'observe',
      handle: () => {
        entered();
        return new Promise(() => {});
      },
    });
    const next = vi.fn();
    hooks.register({ id: 'next', on: 'message:received', mode: 'observe', handle: next });
    const pending = hooks.run('message:received', payload, { ...context, signal: abort.signal });
    const checked = expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    await started;
    abort.abort();
    await checked;
    expect(next).not.toHaveBeenCalled();
  });

  it('validates context providers and enforces a combined context budget', async () => {
    const hooks = new RuntimeHooks();
    hooks.register({
      id: 'oversized',
      on: 'message:received',
      mode: 'guard',
      handle: () => ({
        kind: 'context',
        entries: [
          {
            ref: { id: 'x', kind: 'brief', surface: 'word', title: 'Facts' },
            value: { as: 'text', text: 'x'.repeat(70_000) },
          },
        ],
      }),
    });
    await expect(hooks.run('message:received', payload, context)).rejects.toThrow(
      'Required hook oversized failed',
    );
  });

  it('refuses guards on after-effects and rejects malformed definitions', () => {
    const hooks = new RuntimeHooks();
    expect(() =>
      hooks.register({
        id: 'undo-late',
        on: 'effect:after',
        mode: 'guard',
        handle: () => ({ kind: 'block', reason: 'late' }),
      }),
    ).toThrow('observation-only');
    expect(() =>
      hooks.register({
        id: 'bad',
        on: 'message:received',
        mode: 'observe',
        timeoutMs: Infinity,
        handle: () => {},
      }),
    ).toThrow('timeout');
    expect(() =>
      hooks.register({
        id: 'bad',
        on: 'unknown',
        mode: 'observe',
        handle: () => {},
      } as unknown as RuntimeHook),
    ).toThrow('Invalid runtime hook');
  });

  it('bounds diagnostics and returns defensive copies', async () => {
    const hooks = new RuntimeHooks();
    hooks.register({ id: 'observe', on: 'message:received', mode: 'observe', handle: () => {} });
    for (let i = 0; i < 270; i++) await hooks.run('message:received', payload, context);
    expect(hooks.records()).toHaveLength(256);
    hooks.records()[0]!.hookId = 'corrupt';
    expect(hooks.records()[0]?.hookId).toBe('observe');
  });
});

describe('runtime extension installation', () => {
  it('namespaces hooks and triggers, and closes registration after setup', () => {
    const services = { hooks: new RuntimeHooks(), triggers: new TriggerRegistry() };
    let api!: RuntimeExtensionApi;
    const off = installRuntimeExtensions(
      [
        {
          id: 'tenant',
          setup(a) {
            api = a;
            a.on({ id: 'receive', on: 'message:received', mode: 'observe', handle: () => {} });
            a.trigger({ id: 'mail', on: 'mail-received', handle: () => ({ kind: 'continue' }) });
          },
        },
      ],
      services,
    );
    expect(services.hooks.list()[0]?.id).toBe('tenant/receive');
    expect(services.triggers.size).toBe(1);
    expect(() =>
      api.on({ id: 'late', on: 'message:received', mode: 'observe', handle: () => {} }),
    ).toThrow('closed');
    off();
    off();
    expect(services.hooks.list()).toEqual([]);
    expect(services.triggers.size).toBe(0);
  });

  it('rolls back a partial bundle without removing preexisting registrations', () => {
    const services = { hooks: new RuntimeHooks(), triggers: new TriggerRegistry() };
    services.hooks.register({
      id: 'existing',
      on: 'message:received',
      mode: 'observe',
      handle: () => {},
    });
    expect(() =>
      installRuntimeExtensions(
        [
          {
            id: 'broken',
            setup(api) {
              api.on({ id: 'first', on: 'message:received', mode: 'observe', handle: () => {} });
              throw new Error('setup failed');
            },
          },
        ],
        services,
      ),
    ).toThrow('setup failed');
    expect(services.hooks.list().map((h) => h.id)).toEqual(['existing']);
  });
});
