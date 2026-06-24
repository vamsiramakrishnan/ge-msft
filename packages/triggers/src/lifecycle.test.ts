import { describe, it, expect, vi } from 'vitest';
import { EventBus } from './bus.js';
import { TriggerRegistry, type Trigger } from './registry.js';
import { debounce, type Scheduler } from './debounce.js';
import { coauthorOrigin, eventOrigin, CONTINUE, type HostEvent } from './event.js';

// A deterministic scheduler that records ms and lets the test invoke / clear timers by hand.
function fakeScheduler(): {
  scheduler: Scheduler;
  run: () => void;
  setCalls: number;
  clearCalls: number;
  lastMs: number | null;
  has: () => boolean;
} {
  const holder: { fn: (() => void) | null } = { fn: null };
  const counters = { setCalls: 0, clearCalls: 0, lastMs: null as number | null };
  return {
    scheduler: {
      set: (fn, ms) => {
        counters.setCalls += 1;
        counters.lastMs = ms;
        holder.fn = fn;
        return Symbol('handle');
      },
      clear: () => {
        counters.clearCalls += 1;
        holder.fn = null;
      },
    },
    run: () => holder.fn?.(),
    get setCalls() {
      return counters.setCalls;
    },
    get clearCalls() {
      return counters.clearCalls;
    },
    get lastMs() {
      return counters.lastMs;
    },
    has: () => holder.fn !== null,
  };
}

describe('EventBus — additional behavior', () => {
  it('routes each event only to its own type listeners, not to sibling types', () => {
    const bus = new EventBus();
    const onMail = vi.fn();
    const onMeeting = vi.fn();
    bus.on('mail-received', onMail);
    bus.on('meeting-ended', onMeeting);

    bus.emit({ type: 'mail-received', id: 'm1' });
    expect(onMail).toHaveBeenCalledTimes(1);
    expect(onMeeting).not.toHaveBeenCalled();

    bus.emit({ type: 'meeting-ended', id: 'x' });
    expect(onMail).toHaveBeenCalledTimes(1);
    expect(onMeeting).toHaveBeenCalledTimes(1);
  });

  it('passes the exact emitted event object to listeners', () => {
    const bus = new EventBus();
    const seen: HostEvent[] = [];
    bus.on('selection-changed', (e) => seen.push(e));
    const event: HostEvent = {
      type: 'selection-changed',
      surface: 'word',
      origin: 'local',
      preview: 'hello',
    };
    bus.emit(event);
    expect(seen).toEqual([event]);
    expect(seen[0]).toBe(event);
  });

  it('supports multiple listeners on the same type and unsubscribes only one', () => {
    const bus = new EventBus();
    const a = vi.fn();
    const b = vi.fn();
    const offA = bus.on('document-changed', a);
    bus.on('document-changed', b);

    bus.emit({ type: 'document-changed', surface: 'excel', origin: 'local' });
    expect(a).toHaveBeenCalledTimes(1);
    expect(b).toHaveBeenCalledTimes(1);

    offA();
    bus.emit({ type: 'document-changed', surface: 'excel', origin: 'local' });
    expect(a).toHaveBeenCalledTimes(1); // removed
    expect(b).toHaveBeenCalledTimes(2);
  });

  it('emit is a no-op when there are no listeners for the type or wildcard', () => {
    const bus = new EventBus();
    expect(() => bus.emit({ type: 'mail-send', id: 'm1' })).not.toThrow();
  });

  it('clear() removes every typed and wildcard listener', () => {
    const bus = new EventBus();
    const typed = vi.fn();
    const all = vi.fn();
    bus.on('mail-compose', typed);
    bus.on('*', all);

    bus.clear();

    bus.emit({ type: 'mail-compose', id: 'm1' });
    expect(typed).not.toHaveBeenCalled();
    expect(all).not.toHaveBeenCalled();
  });

  it('a wildcard listener sees events of every type, exactly once each', () => {
    const bus = new EventBus();
    const all = vi.fn();
    bus.on('*', all);
    bus.emit({ type: 'session-start', surface: 'word' });
    bus.emit({ type: 'mail-received', id: 'm1' });
    bus.emit({ type: 'estate-changed', source: 'site', id: 's1' });
    expect(all).toHaveBeenCalledTimes(3);
  });

  it('unsubscribing the same listener twice is harmless', () => {
    const bus = new EventBus();
    const fn = vi.fn();
    const off = bus.on('mail-received', fn);
    off();
    expect(() => off()).not.toThrow();
    bus.emit({ type: 'mail-received', id: 'm1' });
    expect(fn).not.toHaveBeenCalled();
  });
});

describe('coauthorOrigin', () => {
  it('returns remote only when the host explicitly says remote (case-insensitive)', () => {
    expect(coauthorOrigin('Remote')).toBe('remote');
    expect(coauthorOrigin('remote')).toBe('remote');
    expect(coauthorOrigin('REMOTE')).toBe('remote');
  });

  it('treats local, unknown, and non-string sources as local (never mis-tag a real local edit)', () => {
    expect(coauthorOrigin('Local')).toBe('local');
    expect(coauthorOrigin('local')).toBe('local');
    expect(coauthorOrigin('something-else')).toBe('local');
    expect(coauthorOrigin(undefined)).toBe('local');
    expect(coauthorOrigin(null)).toBe('local');
    expect(coauthorOrigin(42)).toBe('local');
    expect(coauthorOrigin({})).toBe('local');
    expect(coauthorOrigin('')).toBe('local');
    // Whitespace is not the literal token => local.
    expect(coauthorOrigin(' remote ')).toBe('local');
  });
});

describe('eventOrigin', () => {
  it('returns the origin for content events that carry one', () => {
    expect(eventOrigin({ type: 'document-changed', surface: 'word', origin: 'remote' })).toBe(
      'remote',
    );
    expect(eventOrigin({ type: 'selection-changed', surface: 'word', origin: 'local' })).toBe(
      'local',
    );
  });

  it('returns undefined for events with no coauthoring origin', () => {
    expect(eventOrigin({ type: 'mail-send', id: 'm1' })).toBeUndefined();
    expect(eventOrigin({ type: 'session-start', surface: 'excel' })).toBeUndefined();
    expect(eventOrigin({ type: 'meeting-ended', id: 'x' })).toBeUndefined();
  });
});

describe('TriggerRegistry — registration lifecycle', () => {
  it('size reflects registrations and unregister via the returned disposer', () => {
    const reg = new TriggerRegistry();
    expect(reg.size).toBe(0);
    const off = reg.register({ id: 'a', on: 'document-changed', handle: () => CONTINUE });
    expect(reg.size).toBe(1);
    off();
    expect(reg.size).toBe(0);
  });

  it('disposer only removes its own trigger and is idempotent', async () => {
    const reg = new TriggerRegistry();
    const h1 = vi.fn(() => ({ kind: 'automate', query: '1' }) as const);
    const h2 = vi.fn(() => ({ kind: 'automate', query: '2' }) as const);
    const off1 = reg.register({ id: 'one', on: 'document-changed', handle: h1 });
    reg.register({ id: 'two', on: 'document-changed', handle: h2 });

    off1();
    off1(); // calling twice must not throw or remove 'two'
    expect(reg.size).toBe(1);

    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    expect(h1).not.toHaveBeenCalled();
    expect(h2).toHaveBeenCalledOnce();
  });

  it('registerAll registers every trigger in order', async () => {
    const reg = new TriggerRegistry();
    const order: string[] = [];
    const mk = (id: string): Trigger => ({
      id,
      on: 'document-changed',
      handle: () => {
        order.push(id);
        return CONTINUE;
      },
    });
    reg.registerAll([mk('a'), mk('b'), mk('c')]);
    expect(reg.size).toBe(3);
    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    expect(order).toEqual(['a', 'b', 'c']);
  });
});

describe('TriggerRegistry.dispatch — matching semantics', () => {
  it('matches when event type is in an array of `on` types', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'suggest', title: 'hit' }) as const);
    reg.register({ id: 'multi', on: ['selection-changed', 'document-changed'], handle });

    await reg.dispatch({ type: 'selection-changed', surface: 'word', origin: 'local' });
    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    await reg.dispatch({ type: 'mail-received', id: 'm' }); // not in the array
    expect(handle).toHaveBeenCalledTimes(2);
  });

  it('collects outcomes from multiple matching triggers, preserving order, dropping continues', async () => {
    const reg = new TriggerRegistry();
    reg.register({
      id: 'first',
      on: 'document-changed',
      handle: () => ({ kind: 'suggest', title: 'A' }),
    });
    reg.register({ id: 'mid', on: 'document-changed', handle: () => CONTINUE });
    reg.register({
      id: 'last',
      on: 'document-changed',
      handle: () => ({ kind: 'automate', query: 'B' }),
    });

    const out = await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'local' });
    expect(out).toEqual([
      { kind: 'suggest', title: 'A' },
      { kind: 'automate', query: 'B' },
    ]);
  });

  it('awaits async handlers and collects their resolved outcomes', async () => {
    const reg = new TriggerRegistry();
    reg.register({
      id: 'async',
      on: 'meeting-ended',
      handle: async () =>
        new Promise((resolve) =>
          setTimeout(() => resolve({ kind: 'automate', query: 'summarize' }), 0),
        ),
    });
    const out = await reg.dispatch({ type: 'meeting-ended', id: 'mtg1' });
    expect(out).toEqual([{ kind: 'automate', query: 'summarize' }]);
  });

  it('does not run a trigger whose match predicate returns false', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'suggest', title: 'x' }) as const);
    reg.register({
      id: 'guarded',
      on: 'selection-changed',
      match: (e) => e.type === 'selection-changed' && e.preview === 'target',
      handle,
    });

    await reg.dispatch({
      type: 'selection-changed',
      surface: 'word',
      origin: 'local',
      preview: 'other',
    });
    expect(handle).not.toHaveBeenCalled();

    await reg.dispatch({
      type: 'selection-changed',
      surface: 'word',
      origin: 'local',
      preview: 'target',
    });
    expect(handle).toHaveBeenCalledOnce();
  });

  it('drops remote content events by default for array-typed triggers too', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'automate', query: 'go' }) as const);
    reg.register({ id: 't', on: ['document-changed', 'selection-changed'], handle });

    await reg.dispatch({ type: 'document-changed', surface: 'word', origin: 'remote' });
    expect(handle).not.toHaveBeenCalled();
  });

  it('events without an origin (e.g. meeting-ended) are never dropped as remote', async () => {
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'automate', query: 'go' }) as const);
    reg.register({ id: 't', on: 'meeting-ended', handle });
    await reg.dispatch({ type: 'meeting-ended', id: 'm1' });
    expect(handle).toHaveBeenCalledOnce();
  });

  it('returns an empty array when no trigger matches the event type', async () => {
    const reg = new TriggerRegistry();
    reg.register({ id: 't', on: 'mail-send', handle: () => ({ kind: 'block', reason: 'no' }) });
    const out = await reg.dispatch({ type: 'session-start', surface: 'word' });
    expect(out).toEqual([]);
  });
});

describe('TriggerRegistry.gate — veto ordering', () => {
  it('returns the FIRST block and short-circuits later gate triggers', async () => {
    const reg = new TriggerRegistry();
    const later = vi.fn(() => ({ kind: 'block', reason: 'second' }) as const);
    reg.register({ id: 'allow', on: 'mail-send', handle: () => CONTINUE });
    reg.register({
      id: 'veto-1',
      on: 'mail-send',
      handle: () => ({ kind: 'block', reason: 'first' }),
    });
    reg.register({ id: 'veto-2', on: 'mail-send', handle: later });

    const outcome = await reg.gate({ type: 'mail-send', id: 'm1' });
    expect(outcome).toEqual({ kind: 'block', reason: 'first' });
    expect(later).not.toHaveBeenCalled(); // short-circuited
  });

  it('treats non-block outcomes (suggest/automate) as non-vetoes and continues', async () => {
    const reg = new TriggerRegistry();
    reg.register({
      id: 'suggester',
      on: 'pre-actuation',
      handle: () => ({ kind: 'suggest', title: 'fyi' }),
    });
    reg.register({
      id: 'automator',
      on: 'pre-actuation',
      handle: () => ({ kind: 'automate', query: 'q' }),
    });
    const request = { changeId: 'c1' } as never;
    const outcome = await reg.gate({ type: 'pre-actuation', request });
    expect(outcome).toEqual(CONTINUE);
  });

  it('respects ignoreRemote in the gate path: a remote event does not trigger the veto', async () => {
    const reg = new TriggerRegistry();
    const veto = vi.fn(() => ({ kind: 'block', reason: 'no' }) as const);
    reg.register({ id: 'veto', on: 'comment-added', handle: veto });

    const outcome = await reg.gate({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'c1',
    });
    expect(outcome).toEqual(CONTINUE);
    expect(veto).not.toHaveBeenCalled();
  });

  it('awaits async gate handlers before deciding', async () => {
    const reg = new TriggerRegistry();
    reg.register({
      id: 'slow-veto',
      on: 'mail-send',
      handle: async () =>
        new Promise((resolve) => setTimeout(() => resolve({ kind: 'block', reason: 'slow' }), 0)),
    });
    const outcome = await reg.gate({ type: 'mail-send', id: 'm1' });
    expect(outcome).toEqual({ kind: 'block', reason: 'slow' });
  });

  it('continues when there are no triggers at all', async () => {
    const reg = new TriggerRegistry();
    expect(await reg.gate({ type: 'mail-send', id: 'm1' })).toEqual(CONTINUE);
  });
});

describe('debounce — timer edges', () => {
  it('schedules with the configured delay and clears the prior timer on each new call', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 200, s.scheduler);

    d('a');
    expect(s.setCalls).toBe(1);
    expect(s.clearCalls).toBe(0);
    expect(s.lastMs).toBe(200);

    d('b'); // clears previous, sets again
    expect(s.clearCalls).toBe(1);
    expect(s.setCalls).toBe(2);
  });

  it('flush invokes the pending call immediately with the latest args and clears the timer', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 500, s.scheduler);
    d('first');
    d('second');
    d.flush();
    expect(fn).toHaveBeenCalledTimes(1);
    expect(fn).toHaveBeenCalledWith('second');
    expect(s.clearCalls).toBe(2); // one per re-call (1) + one in flush (1)
    // After flush there is no pending timer, so running it again does nothing.
    expect(s.has()).toBe(false);
  });

  it('flush with nothing pending is a no-op', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 100, s.scheduler);
    d.flush();
    expect(fn).not.toHaveBeenCalled();
    expect(s.clearCalls).toBe(0);
  });

  it('after a fired call, a fresh call starts a new debounce window', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 100, s.scheduler);
    d('one');
    s.run();
    expect(fn).toHaveBeenCalledWith('one');

    d('two');
    s.run();
    expect(fn).toHaveBeenCalledTimes(2);
    expect(fn).toHaveBeenLastCalledWith('two');
  });

  it('flush after the timer already fired does nothing (no double-invoke)', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 100, s.scheduler);
    d('x');
    s.run(); // timer fires, pending cleared, handle nulled
    d.flush();
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('cancel after a flush is harmless and keeps the single invocation', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 100, s.scheduler);
    d('y');
    d.flush();
    d.cancel();
    expect(fn).toHaveBeenCalledTimes(1);
  });

  it('preserves multiple arguments through to the trailing invocation', () => {
    const s = fakeScheduler();
    const fn = vi.fn();
    const d = debounce(fn, 10, s.scheduler);
    d(1, 'two', { three: true });
    s.run();
    expect(fn).toHaveBeenCalledWith(1, 'two', { three: true });
  });

  it('uses the default (real timer) scheduler when none is injected', async () => {
    const fn = vi.fn();
    const d = debounce(fn, 1);
    d('real');
    expect(fn).not.toHaveBeenCalled();
    await new Promise((r) => setTimeout(r, 10));
    expect(fn).toHaveBeenCalledWith('real');
    d.cancel();
  });
});

describe('integration — bus + debounce + registry (event source pipeline)', () => {
  it('coalesces high-frequency selection events, then dispatches once through the registry', async () => {
    const s = fakeScheduler();
    const bus = new EventBus();
    const reg = new TriggerRegistry();
    const handle = vi.fn(() => ({ kind: 'suggest', title: 'ambient' }) as const);
    reg.register({ id: 'ambient', on: 'selection-changed', handle });

    // Event source: debounce raw selection moves, dispatch the survivor.
    const dispatched: HostEvent[] = [];
    const onSelection = debounce(
      (e: HostEvent) => {
        dispatched.push(e);
        void reg.dispatch(e);
      },
      150,
      s.scheduler,
    );
    bus.on('selection-changed', onSelection);

    bus.emit({ type: 'selection-changed', surface: 'word', origin: 'local', preview: '1' });
    bus.emit({ type: 'selection-changed', surface: 'word', origin: 'local', preview: '2' });
    bus.emit({ type: 'selection-changed', surface: 'word', origin: 'local', preview: '3' });

    expect(dispatched).toHaveLength(0); // nothing fired yet
    s.run();
    await Promise.resolve();

    expect(dispatched).toHaveLength(1);
    expect((dispatched[0] as { preview?: string }).preview).toBe('3');
    expect(handle).toHaveBeenCalledOnce();
  });
});
