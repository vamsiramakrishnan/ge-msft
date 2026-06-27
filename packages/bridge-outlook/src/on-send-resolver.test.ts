import { afterEach, describe, expect, it } from 'vitest';
import { TriggerRegistry } from '@ge/triggers';
import {
  activeItemIdResolver,
  createMessageSendHandler,
  type OnSendCompletedOptions,
} from './on-send.js';

/**
 * Covers the Office.js-touching parts of the on-send glue that the pure `events.test.ts` cannot
 * reach: `activeItemIdResolver` reading `Office.context.mailbox.item.itemId`, and the
 * `opts.resolveItemId` wiring in {@link createMessageSendHandler}. A tiny fake `globalThis.Office`
 * stands in for the host; the resolver is defensive (Office may be undefined / throw).
 */

interface Installed {
  restore(): void;
}

function installOffice(item: unknown, opts: { throws?: boolean } = {}): Installed {
  const office = {
    context: {
      get mailbox(): unknown {
        if (opts.throws) throw new Error('mailbox access exploded');
        return { item };
      },
    },
  };
  const g = globalThis as unknown as Record<string, unknown>;
  const prev = g.Office;
  g.Office = office;
  return {
    restore() {
      g.Office = prev;
    },
  };
}

let active: Installed | undefined;
afterEach(() => {
  active?.restore();
  active = undefined;
});

describe('activeItemIdResolver', () => {
  it('returns the saved item id when a read-mode item is active', () => {
    active = installOffice({ itemId: 'AAMk-saved' });
    expect(activeItemIdResolver()).toBe('AAMk-saved');
  });

  it('returns undefined for an unsaved draft (no itemId field)', () => {
    active = installOffice({ subject: 'Draft' });
    expect(activeItemIdResolver()).toBeUndefined();
  });

  it('returns undefined when itemId is an empty string', () => {
    active = installOffice({ itemId: '' });
    expect(activeItemIdResolver()).toBeUndefined();
  });

  it('returns undefined when itemId is a non-string value', () => {
    active = installOffice({ itemId: 12345 });
    expect(activeItemIdResolver()).toBeUndefined();
  });

  it('returns undefined when there is no active item', () => {
    active = installOffice(undefined);
    expect(activeItemIdResolver()).toBeUndefined();
  });

  it('returns undefined defensively when accessing the mailbox throws', () => {
    active = installOffice({ itemId: 'x' }, { throws: true });
    expect(activeItemIdResolver()).toBeUndefined();
  });

  it('returns undefined when Office is not present at all', () => {
    const g = globalThis as unknown as Record<string, unknown>;
    const prev = g.Office;
    g.Office = undefined;
    try {
      expect(activeItemIdResolver()).toBeUndefined();
    } finally {
      g.Office = prev;
    }
  });
});

describe('createMessageSendHandler with the default activeItemIdResolver', () => {
  it('threads the real resolver output into the gated mail-send event', async () => {
    active = installOffice({ itemId: 'AAMk-live' });
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
    const handler = createMessageSendHandler(registry, { resolveItemId: activeItemIdResolver });
    const calls: OnSendCompletedOptions[] = [];
    await handler({ completed: (o?: OnSendCompletedOptions) => calls.push(o ?? {}) });
    expect(seenId).toBe('AAMk-live');
    expect(calls).toEqual([{ allowEvent: true }]);
  });
});
