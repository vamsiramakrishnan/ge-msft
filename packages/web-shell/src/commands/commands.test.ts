import { describe, it, expect, vi } from 'vitest';
import { buildAskSelectionSeed, askSelection, ASK_SELECTION_SEED_KEY } from './commands.js';

/**
 * Tests for the right-click "Ask Gemini about this" function-command. The load-bearing behavior:
 * `buildAskSelectionSeed` grounds the selection as an `@this` `assist` seed (and stays valid on an
 * empty selection), and `askSelection` reads the host selection, stashes that seed where the pane
 * picks it up, reveals the pane, and ALWAYS completes the Office event — even when the read, the
 * stash, or the reveal throws — so the host command never hangs.
 */

function fakeOffice(
  opts: {
    selection?: string;
    failRead?: boolean;
    showAsTaskpane?: ReturnType<typeof vi.fn>;
  } = {},
) {
  return {
    CoercionType: { Text: 'text' },
    context: {
      document: {
        getSelectedDataAsync: (
          _coercion: unknown,
          cb: (r: { status?: unknown; value?: unknown }) => void,
        ) => {
          if (opts.failRead) {
            cb({ status: 'failed' });
            return;
          }
          cb({ status: 'succeeded', value: opts.selection ?? '' });
        },
      },
    },
    addin: {
      showAsTaskpane: (opts.showAsTaskpane ??
        vi.fn().mockResolvedValue(undefined)) as () => Promise<unknown>,
    },
  };
}

describe('buildAskSelectionSeed', () => {
  it('records a selection without persisting the selected text', () => {
    const seed = buildAskSelectionSeed('The vendor SLA is 99.5%.');
    expect(seed.kind).toBe('ask-selection');
    expect(seed.hasSelection).toBe(true);
    // The raw selected text must NOT ride in the seed (it crosses localStorage).
    expect(JSON.stringify(seed)).not.toContain('99.5');
  });

  it('flags an empty selection', () => {
    const seed = buildAskSelectionSeed('   ');
    expect(seed.kind).toBe('ask-selection');
    expect(seed.hasSelection).toBe(false);
  });
});

describe('askSelection', () => {
  it('reads the selection, stashes the seed, reveals the pane, and completes', async () => {
    const completed = vi.fn();
    const setItem = vi.fn();
    const showAsTaskpane = vi.fn().mockResolvedValue(undefined);
    const office = fakeOffice({ selection: 'renewal in March', showAsTaskpane });

    await askSelection({ completed }, { office, sink: { setItem } });

    expect(setItem).toHaveBeenCalledTimes(1);
    const [key, json] = setItem.mock.calls[0]!;
    expect(key).toBe(ASK_SELECTION_SEED_KEY);
    const seed = JSON.parse(json as string);
    expect(seed).toEqual({ kind: 'ask-selection', hasSelection: true });
    expect(json).not.toContain('renewal'); // the selected text never crosses the channel
    expect(showAsTaskpane).toHaveBeenCalledTimes(1);
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it('still seeds (empty) and completes when the read fails', async () => {
    const completed = vi.fn();
    const setItem = vi.fn();
    const office = fakeOffice({ failRead: true });

    await askSelection({ completed }, { office, sink: { setItem } });

    const seed = JSON.parse(setItem.mock.calls[0]![1] as string);
    expect(seed).toEqual({ kind: 'ask-selection', hasSelection: false });
    expect(completed).toHaveBeenCalledTimes(1);
  });

  it('completes even when the reveal rejects', async () => {
    const completed = vi.fn();
    const setItem = vi.fn();
    const showAsTaskpane = vi.fn().mockRejectedValue(new Error('pane already open'));
    const office = fakeOffice({ selection: 'x', showAsTaskpane });

    await askSelection({ completed }, { office, sink: { setItem } });

    expect(completed).toHaveBeenCalledTimes(1);
  });

  it('completes even when the stash throws (storage blocked)', async () => {
    const completed = vi.fn();
    const setItem = vi.fn(() => {
      throw new Error('quota exceeded');
    });
    const office = fakeOffice({ selection: 'x' });

    await askSelection({ completed }, { office, sink: { setItem } });

    expect(completed).toHaveBeenCalledTimes(1);
  });

  it('degrades to an empty seed and completes when no selection API is present', async () => {
    const completed = vi.fn();
    const setItem = vi.fn();

    await askSelection({ completed }, { office: {}, sink: { setItem } });

    const seed = JSON.parse(setItem.mock.calls[0]![1] as string);
    expect(seed).toEqual({ kind: 'ask-selection', hasSelection: false });
    expect(completed).toHaveBeenCalledTimes(1);
  });
});
