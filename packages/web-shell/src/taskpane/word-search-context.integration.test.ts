// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { scriptedClient, mountStack, type MountedStack } from '../test-harness/index.js';
import { installFakeWord, wordSeed, type WordSimulator } from '../test-harness/fake-word.js';

/**
 * docs/COMMAND-RELIABILITY.md: on a long document a search for a heading returned only the heading
 * line, so a model asked to comment on the sentences under it searched until it ran out of turns.
 * A hit now carries the paragraphs that follow it.
 */
let sim: WordSimulator | undefined;
let ui: MountedStack | undefined;
afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

describe('Word search context', () => {
  it('returns the paragraphs after a matched heading', async () => {
    sim = installFakeWord(
      wordSeed({
        paragraphs: [
          { text: 'Payment terms' },
          { text: 'Payment is due within 45 days of the invoice date.' },
          { text: 'The fee will be an amount agreed later.' },
          { text: 'Signed by both parties.' },
        ],
      } as never),
    );
    ui = mountStack({ surface: 'word', client: scriptedClient([]) });
    await ui.flush();
    const [hit] = await ui.bridge.searchDocument!('Payment terms');
    const text = hit?.value.as === 'text' ? hit.value.text : '';
    expect(text).toContain('Payment is due within 45 days');
    expect(text).toContain('amount agreed later');
    expect(text).not.toContain('Signed by both parties'); // bounded to two following paragraphs
  });
});
