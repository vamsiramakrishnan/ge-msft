import { afterEach, describe, expect, it } from 'vitest';
import { installFakeWord, wordSeed, type WordSimulator } from './fake-word.js';

let simulator: WordSimulator | undefined;
afterEach(() => {
  simulator?.restore();
  simulator = undefined;
});
describe('Word simulator current-revision readback', () => {
  it('reads actual current content after native and manual changes, without copying the requested replacement', async () => {
    simulator = installFakeWord(
      wordSeed({ paragraphs: [{ text: 'Prefix old suffix.', styleBuiltIn: 'Normal' }] }),
    );
    await Word.run(async (context) => {
      const hits = context.document.body.search('old', { matchCase: true });
      hits.load('items/text');
      await context.sync();
      const range = hits.items[0]!;
      const observed = range.getReviewedText('Current');
      expect(observed.value).toBe('old');
      range.insertText('generated', Word.InsertLocation.replace);
      await context.sync();
      expect(observed.value).toBe('generated');
      simulator!.seed.paragraphs[0]!.text = 'Prefix manual suffix.';
      expect(observed.value).toBe('manual');
    });
  });
  it('refuses a drifted simulated span rather than reporting a requested value as verified', async () => {
    simulator = installFakeWord(
      wordSeed({ paragraphs: [{ text: 'Prefix old suffix.', styleBuiltIn: 'Normal' }] }),
    );
    await Word.run(async (context) => {
      const hits = context.document.body.search('old', { matchCase: true });
      hits.load('items/text');
      await context.sync();
      const observed = hits.items[0]!.getReviewedText('Current');
      simulator!.seed.paragraphs[0]!.text = 'A different prefix old suffix.';
      expect(() => observed.value).toThrow('drifted');
    });
  });
});
