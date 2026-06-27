import { describe, it, expect } from 'vitest';
import { TriggerRegistry, type TriggerOutcome } from '@ge/triggers';
import { composeEvent, receivedEvent, sendEvent, decideSend } from './events.js';
import {
  createMessageSendHandler,
  type OnSendCompletedOptions,
  type OnSendEvent,
} from './on-send.js';

describe('outlook event builders (pure)', () => {
  it('builds a mail-compose without an id', () => {
    expect(composeEvent()).toEqual({ type: 'mail-compose' });
  });

  it('builds a mail-compose with an id when given one', () => {
    expect(composeEvent('AAMk-1')).toEqual({ type: 'mail-compose', id: 'AAMk-1' });
  });

  it('builds a mail-received with its required id', () => {
    expect(receivedEvent('AAMk-2')).toEqual({ type: 'mail-received', id: 'AAMk-2' });
  });

  it('builds a mail-send with and without an id', () => {
    expect(sendEvent()).toEqual({ type: 'mail-send' });
    expect(sendEvent('AAMk-3')).toEqual({ type: 'mail-send', id: 'AAMk-3' });
  });
});

describe('decideSend (pure on-send decision)', () => {
  it('blocks with the reason as message on a block outcome', () => {
    const outcome: TriggerOutcome = {
      kind: 'block',
      reason: 'External recipient on a confidential thread',
    };
    expect(decideSend(outcome)).toEqual({
      allowEvent: false,
      message: 'External recipient on a confidential thread',
    });
  });

  it('allows on continue', () => {
    expect(decideSend({ kind: 'continue' })).toEqual({ allowEvent: true });
  });

  it('allows on suggest', () => {
    expect(decideSend({ kind: 'suggest', title: 'Consider citing the SLA' })).toEqual({
      allowEvent: true,
    });
  });

  it('allows on automate', () => {
    expect(decideSend({ kind: 'automate', query: 'summarize thread' })).toEqual({
      allowEvent: true,
    });
  });
});

describe('createMessageSendHandler (host glue with a real registry, stubbed event)', () => {
  function fakeEvent() {
    const calls: OnSendCompletedOptions[] = [];
    return {
      calls,
      event: { completed: (options?: OnSendCompletedOptions) => calls.push(options ?? {}) },
    };
  }

  it('allows send when no trigger blocks', async () => {
    const registry = new TriggerRegistry();
    const handler = createMessageSendHandler(registry);
    const { calls, event } = fakeEvent();
    await handler(event);
    expect(calls).toEqual([{ allowEvent: true }]);
  });

  it('blocks send and surfaces the reason via errorMessage when a trigger blocks', async () => {
    const registry = new TriggerRegistry();
    registry.register({
      id: 'block-external',
      on: 'mail-send',
      handle: () => ({ kind: 'block', reason: 'Confidential — review before sending externally' }),
    });
    const handler = createMessageSendHandler(registry);
    const { calls, event } = fakeEvent();
    await handler(event);
    expect(calls).toEqual([
      { allowEvent: false, errorMessage: 'Confidential — review before sending externally' },
    ]);
  });

  it('passes the resolved item id into the gated mail-send event', async () => {
    const registry = new TriggerRegistry();
    let seenId: string | undefined = 'unset';
    registry.register({
      id: 'capture-id',
      on: 'mail-send',
      handle: (e) => {
        seenId = e.type === 'mail-send' ? e.id : undefined;
        return { kind: 'continue' };
      },
    });
    const handler = createMessageSendHandler(registry, { resolveItemId: () => 'AAMk-send' });
    const { event } = fakeEvent();
    await handler(event);
    expect(seenId).toBe('AAMk-send');
  });

  it('fails safe (allowEvent true) if the gate throws', async () => {
    const registry = new TriggerRegistry();
    registry.register({
      id: 'boom',
      on: 'mail-send',
      handle: () => {
        throw new Error('trigger exploded');
      },
    });
    const handler = createMessageSendHandler(registry);
    const { calls, event } = fakeEvent();
    await handler(event);
    expect(calls).toEqual([{ allowEvent: true }]);
  });

  // HIGH-2 regression: a decided block whose first `completed(...)` throws must NOT be downgraded
  // to a second `completed({ allowEvent: true })`. The block must stand (the error propagates),
  // never silently letting the mail send.
  it('does NOT downgrade a real block to allow if completed() throws on the block call', async () => {
    const registry = new TriggerRegistry();
    registry.register({
      id: 'block-throwing-host',
      on: 'mail-send',
      handle: () => ({ kind: 'block', reason: 'Confidential — do not send externally' }),
    });
    const handler = createMessageSendHandler(registry);

    const calls: OnSendCompletedOptions[] = [];
    const event: OnSendEvent = {
      completed: (options?: OnSendCompletedOptions) => {
        calls.push(options ?? {});
        // The host's first completed() for the genuine block throws.
        throw new Error('host completed() exploded');
      },
    };

    await expect(handler(event)).rejects.toThrow('host completed() exploded');
    // Exactly one completed call, and it stayed a block — never a second allowEvent:true.
    expect(calls).toEqual([
      { allowEvent: false, errorMessage: 'Confidential — do not send externally' },
    ]);
    expect(calls.some((c) => c.allowEvent === true)).toBe(false);
  });
});
