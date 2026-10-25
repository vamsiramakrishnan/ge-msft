import { afterEach, describe, expect, it } from 'vitest';
import { WordBridge } from '@ge/bridge-word';
import type { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import { installFakeWord, wordSeed, type WordSimulator } from './fake-word.js';

let sim: WordSimulator | undefined;
afterEach(() => {
  sim?.restore();
  sim = undefined;
});
function install() {
  sim = installFakeWord(
    wordSeed({
      selectionText: 'Middle',
      paragraphs: [
        { text: 'First paragraph.', styleBuiltIn: 'Normal' },
        { text: 'Middle paragraph.', styleBuiltIn: 'Normal' },
        { text: 'Last paragraph.', styleBuiltIn: 'Normal' },
      ],
    }),
  );
  return new AssistSession(new WordBridge(), {} as StreamAssistClient, {
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
    autoAttach: [],
  });
}
async function run(program: string) {
  const session = install();
  const events = [];
  for await (const event of session.runCommandProgram(program, { approvePlan: () => true }))
    events.push(event);
  return events;
}
describe('Word insertion positions through the real compiler/bridge (native simulator)', () => {
  it('appends at the explicit document end without replacing the current selection', async () => {
    const events = await run('/insert-text position=end text=" APPENDED"');
    expect(events).toContainEqual(
      expect.objectContaining({
        type: 'write-result',
        result: expect.objectContaining({ ok: true, location: 'body:end' }),
      }),
    );
    expect(sim!.snapshot().bodyText).toBe(
      'First paragraph.\nMiddle paragraph.\nLast paragraph. APPENDED',
    );
    expect(sim!.seed.selectionText).toBe('Middle');
  });
  it('prepends at the explicit document start while retaining the selected paragraph', async () => {
    await run('/insert-text position=start text="PREPENDED "');
    expect(sim!.snapshot().bodyText).toBe(
      'PREPENDED First paragraph.\nMiddle paragraph.\nLast paragraph.',
    );
  });
  it('keeps an explicit end scoped to a supplied content anchor', async () => {
    await run('/insert-text match="Middle" position=end text=" SCOPED"');
    expect(sim!.snapshot().bodyText).toBe(
      'First paragraph.\nMiddle SCOPED paragraph.\nLast paragraph.',
    );
  });
  it('preserves the existing selection replacement default when no position is requested', async () => {
    await run('/insert-text text="Replacement"');
    expect(sim!.snapshot().bodyText).toBe(
      'First paragraph.\nReplacement paragraph.\nLast paragraph.',
    );
  });
  it('rejects an invalid position before any native mutation', async () => {
    const events = await run('/insert-text position=somewhere text="BAD"');
    expect(sim!.snapshot().inserts).toHaveLength(0);
    expect(events.some((event) => event.type === 'command' && 'error' in event.compiled)).toBe(
      true,
    );
  });
});
