import { describe, it, expect, vi } from 'vitest';
import { EventBus } from './bus.js';
import { TriggerRegistry } from './registry.js';
import { debounce, type Scheduler } from './debounce.js';
import type { HostEvent } from './event.js';

describe('EventBus', () => {
  it('delivers to type listeners and wildcard, and unsubscribes', () => {
    const bus = new EventBus();
    const typed = vi.fn();
    const all = vi.fn();
    const off = bus.on('mail-received', typed);
    bus.on('*', all);
    bus.emit({ type: 'mail-received', id: 'm1' });
    expect(typed).toHaveBeenCalledOnce();
    expect(all).toHaveBeenCalledOnce();
    off();
    bus.emit({ type: 'mail-received', id: 'm2' });
    expect(typed).toHaveBeenCalledOnce(); // not called after unsubscribe
    expect(all).toHaveBeenCalledTimes(2);
  });
});

describe('TriggerRegistry.dispatch', () => {
  it('fires matching triggers and collects non-continue outcomes', async () => {
    const reg = new TriggerRegistry();
    reg.register({
      id: 'ambient-selection',
      on: 'selection-changed',
      handle: (e) => ({ kind: 'suggest', title: `look at ${e.type}` }),
    });
    reg.register({ id: 'noop', on: 'document-changed', handle: () => ({ kind: 'continue' }) });

    const out = await reg.dispatch({ type: 'selection-changed', surface: 'word', origin: 'local' });
    expect(out).toEqual([{ kind: 'suggest', title: 'look at selection-changed' }]);

    const none = await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    expect(none).toEqual([]);
  });

  it('ignores remote (coauthor / own-write) events by default', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'automate', query: 'go' }) as const);
    reg.register({ id: 't', on: 'document-changed', handle });

    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'remote' });
    expect(handle).not.toHaveBeenCalled();

    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    expect(handle).toHaveBeenCalledOnce();
  });

  it('honors ignoreRemote:false and a match predicate', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'automate', query: 'x' }) as const);
    reg.register({
      id: 'remote-ok',
      on: 'comment-added',
      ignoreRemote: false,
      match: (e) => e.type === 'comment-added' && e.text === 'please fix',
      handle,
    });
    await reg.dispatch({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'c1',
      text: 'other',
    });
    expect(handle).not.toHaveBeenCalled();
    await reg.dispatch({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'c1',
      text: 'please fix',
    });
    expect(handle).toHaveBeenCalledOnce();
  });
});

describe('TriggerRegistry.gate (PreToolUse analog)', () => {
  const onSend: HostEvent = { type: 'mail-send', id: 'm1' };

  it('blocks when a gate trigger vetoes, returning its reason', async () => {
    const reg = new TriggerRegistry();
    reg.register({ id: 'allow', on: 'mail-send', handle: () => ({ kind: 'continue' }) });
    reg.register({
      id: 'ungrounded-check',
      on: 'mail-send',
      handle: () => ({ kind: 'block', reason: 'A claim in the draft is not grounded.' }),
    });
    const outcome = await reg.gate(onSend);
    expect(outcome).toEqual({ kind: 'block', reason: 'A claim in the draft is not grounded.' });
  });

  it('continues when no gate blocks', async () => {
    const reg = new TriggerRegistry();
    reg.register({ id: 'allow', on: 'mail-send', handle: () => ({ kind: 'continue' }) });
    expect(await reg.gate(onSend)).toEqual({ kind: 'continue' });
  });
});

describe('debounce', () => {
  // Deterministic scheduler: holds the single pending timer callback.
  function fakeScheduler(): { scheduler: Scheduler; run: () => void } {
    const holder: { fn: (() => void) | null } = { fn: null };
    return {
      scheduler: {
        set: (fn) => {
          holder.fn = fn;
          return 1;
        },
        clear: () => {
          holder.fn = null;
        },
      },
      run: () => holder.fn?.(),
    };
  }

  it('coalesces rapid calls to a single trailing invocation', () => {
    const { scheduler, run } = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 200, scheduler);
    d('a');
    d('b');
    d('c'); // only the last survives
    expect(fn).not.toHaveBeenCalled();
    run();
    expect(fn).toHaveBeenCalledTimes(1);
    expect(fn).toHaveBeenCalledWith('c');
  });

  it('cancel drops the pending call', () => {
    const { scheduler, run } = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 50, scheduler);
    d('x');
    d.cancel();
    run();
    expect(fn).not.toHaveBeenCalled();
  });
});
