import { describe, expect, it } from 'vitest';
import { wordEdits, type WordEdit } from './actuate-plan.js';

/** Apply edits to the old words, the way the host does (back to front). */
function apply(oldText: string, edits: WordEdit[]): string {
  const words = oldText.split(/\s+/).filter(Boolean);
  for (const e of [...edits].reverse())
    words.splice(e.start, e.end - e.start, ...(e.text ? [e.text] : []));
  return words.join(' ');
}

describe('wordEdits', () => {
  it('marks only the changed words of a grammar fix (UAT #17)', () => {
    const before =
      'Seluruh pengembangan solusi AI wajib mengikuti regulasi yang berlaku, tidak terbatas pada POJK yang diperbarui';
    const after =
      'Seluruh pengembangan solusi AI wajib mengikuti regulasi yang berlaku, termasuk POJK yang diperbarui';
    const edits = wordEdits(before, after)!;
    expect(edits).toEqual([{ start: 9, end: 12, text: 'termasuk' }]);
    expect(apply(before, edits)).toBe(after);
  });

  it('widens a pure insertion or deletion to a neighbour word, and round-trips', () => {
    const cases: Array<[string, string]> = [
      ['the quick fox jumps', 'the quick brown fox jumps'],
      ['the quick brown fox jumps', 'the quick fox jumps'],
      ['a b c d e', 'a x c d y'],
      ['one two three four', 'one two three four five'],
      ['one two three four', 'zero one two three four'],
    ];
    for (const [before, after] of cases) {
      const edits = wordEdits(before, after)!;
      expect(edits.every((e) => e.end > e.start)).toBe(true);
      expect(apply(before, edits)).toBe(after);
    }
  });

  it('gives up when nearly everything changed', () => {
    expect(wordEdits('alpha beta gamma delta', 'one two three four')).toBeUndefined();
  });
});
