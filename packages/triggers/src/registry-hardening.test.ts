import { afterEach, describe, expect, it, vi } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { TriggerRegistry } from './registry.js';
import type { HostEvent } from './event.js';

afterEach(() => {
  vi.useRealTimers();
});
describe('bounded, isolated trigger dispatch', () => {
  it('times out a required check, cancels its work, and returns a block', async () => {
    vi.useFakeTimers();
    const registry = new TriggerRegistry();
    let signal!: AbortSignal;
    registry.register({
      id: 'hung',
      on: 'mail-send',
      timeoutMs: 15,
      handle: (_, context) => {
        signal = context.signal;
        return new Promise(() => {});
      },
    });
    const pending = registry.gate({ type: 'mail-send' });
    await vi.advanceTimersByTimeAsync(16);
    expect(await pending).toMatchObject({ kind: 'block' });
    expect(signal.aborted).toBe(true);
  });
  it('isolates observer and matcher failures and reports metadata without exception bodies', async () => {
    const registry = new TriggerRegistry();
    const diagnostic = vi.fn();
    registry.subscribe(diagnostic);
    registry.register({
      id: 'broken',
      on: 'mail-received',
      match: () => {
        throw new Error('private content');
      },
      handle: () => ({ kind: 'continue' }),
    });
    registry.register({
      id: 'next',
      on: 'mail-received',
      handle: () => ({ kind: 'suggest', title: 'Review' }),
    });
    expect(await registry.dispatch({ type: 'mail-received', id: 'sensitive-id' })).toEqual([
      { kind: 'suggest', title: 'Review' },
    ]);
    expect(diagnostic).toHaveBeenCalledWith({
      triggerId: 'broken',
      event: 'mail-received',
      outcome: 'failed',
    });
    expect(JSON.stringify(diagnostic.mock.calls)).not.toContain('private content');
  });
  it('prevents a trigger from mutating the approved effect request', async () => {
    const registry = new TriggerRegistry();
    registry.register({
      id: 'mutate',
      on: 'pre-actuation',
      handle: (event) => {
        if (event.type === 'pre-actuation') event.request.params.text = 'evil';
        return { kind: 'continue' };
      },
    });
    const event: HostEvent = {
      type: 'pre-actuation',
      request: {
        changeId: asChangeId('c'),
        surface: 'word',
        kind: 'tracked-change',
        params: { text: 'approved' },
      },
    };
    expect(await registry.gate(event)).toMatchObject({ kind: 'block' });
    expect(event.request.params.text).toBe('approved');
  });
  it('uses a registration snapshot so self-removal cannot skip the next handler', async () => {
    const registry = new TriggerRegistry();
    const off = registry.register({
      id: 'once',
      on: 'mail-received',
      priority: 1,
      handle: () => {
        off();
        return { kind: 'continue' };
      },
    });
    registry.register({
      id: 'next',
      on: 'mail-received',
      handle: () => ({ kind: 'suggest', title: 'Still runs' }),
    });
    expect(await registry.dispatch({ type: 'mail-received', id: 'x' })).toEqual([
      { kind: 'suggest', title: 'Still runs' },
    ]);
  });
});
