import { afterEach, describe, expect, it } from 'vitest';
import { asChangeId, type ActuationRequest, type ContextRef } from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { WordBridge } from '@ge/bridge-word';
import { installFakeWord, wordSeed, type WordSimulator } from './fake-word.js';

/**
 * C2 — the REAL `WordBridge` + `OfficeWordHost` against the fake Word host: critique annotations
 * (WordApi 1.7), the comment fallback below it, the minted-id inverse + readback, annotation
 * events through `watch()`, and the WordApi 1.8 highlight preview on reveal.
 */
let simulator: WordSimulator | undefined;
afterEach(() => {
  simulator?.restore();
  simulator = undefined;
});

const seed = () =>
  wordSeed({
    paragraphs: [
      { text: 'Master Services Agreement', styleBuiltIn: 'Heading1' },
      { text: 'The SLA is 99.5%. Credits apply beyond 99.5% downtime.', styleBuiltIn: 'Normal' },
    ],
  });

function request(params: ActuationRequest['params'], id = 'chg-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-annotation', surface: 'word', params };
}

const clauseRef: ContextRef = {
  id: 'ctx:sla',
  kind: 'paragraph',
  surface: 'word',
  title: 'SLA clause',
  anchor: { matchText: '99.5%', contextHint: 'Credits apply' },
};

describe('Word critique annotations against the simulator', () => {
  it('mints a critique addressed inside the hit paragraph, verified by readback, deletable by id', async () => {
    simulator = installFakeWord(seed(), { WordApi: 8 });
    const bridge = new WordBridge();
    const res = await bridge.actuate(
      request({
        target: { matchText: '99.5%', contextHint: 'Credits apply' },
        text: 'Policy requires 99.9%.',
        annotation: { suggestion: '99.9%', severity: 'warning' },
      }),
    );
    expect(res).toMatchObject({
      ok: true,
      location: 'annotation:sim-word-annotation-1',
      completion: 'verified',
      inverse: { op: 'delete-object', objectType: 'annotation', name: 'sim-word-annotation-1' },
    });
    const [annotation] = simulator.snapshot().annotations;
    // The context hint picks the SECOND hit ("beyond 99.5% downtime"); offsets are paragraph-local.
    expect(annotation).toMatchObject({
      paragraphIndex: 1,
      anchor: '99.5%',
      colorScheme: 'Lavender',
      state: 'Created',
    });
    expect(annotation!.start).toBe('The SLA is 99.5%. Credits apply beyond '.length);
    expect(annotation!.length).toBe('99.5%'.length);
    // Popup options need manifest resource ids: none configured ⇒ no popup, no suggestions.
    expect(annotation!.suggestions).toEqual([]);
    // Document content is untouched: a critique is an overlay, not an edit.
    expect(simulator.snapshot().bodyText).toBe(
      'Master Services Agreement\nThe SLA is 99.5%. Credits apply beyond 99.5% downtime.',
    );
    expect(simulator.snapshot().addedComments).toEqual([]);

    // The inverse: delete exactly the minted annotation; a second delete reports failure.
    await Word.run(async (ctx) => {
      ctx.document.getAnnotationById('sim-word-annotation-1').delete();
      await ctx.sync();
    });
    expect(simulator.snapshot().annotations).toEqual([]);
    await expect(
      Word.run(async (ctx) => {
        ctx.document.getAnnotationById('sim-word-annotation-1').load('state');
        await ctx.sync();
      }),
    ).rejects.toThrow('ItemNotFound');
  });

  it('falls back to an anchored comment on a WordApi 1.6 host', async () => {
    simulator = installFakeWord(seed(), { WordApi: 6 });
    const res = await new WordBridge().actuate(
      request({
        target: { matchText: 'Master Services Agreement' },
        text: 'Title should name the parties.',
        annotation: { suggestion: 'Master Services Agreement between Acme and Globex' },
      }),
    );
    expect(res).toMatchObject({ ok: true, location: 'comment', degraded: true });
    expect(simulator.snapshot().annotations).toEqual([]);
    expect(simulator.snapshot().addedComments).toEqual([
      {
        anchor: 'Master Services Agreement',
        text: 'Title should name the parties. Suggestion: Master Services Agreement between Acme and Globex',
      },
    ]);
  });

  it('degrades a drifted anchor to a panel item without writing anything', async () => {
    simulator = installFakeWord(seed(), { WordApi: 8 });
    const res = await new WordBridge().actuate(
      request({ target: { matchText: 'net 90' }, text: 'finding' }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'anchor_drift' } });
    expect(simulator.snapshot().annotations).toEqual([]);
    expect(simulator.snapshot().addedComments).toEqual([]);
  });

  it('streams annotation events through watch() on 1.8, none below 1.7, and unsubscribes', async () => {
    simulator = installFakeWord(seed(), { WordApi: 8 });
    const events: HostEvent[] = [];
    const unsub = new WordBridge().watch((e) => events.push(e));
    await Promise.resolve(); // let the Word.run registration settle
    expect(simulator.annotationHandlerCount('onAnnotationClicked')).toBe(1);
    expect(simulator.annotationHandlerCount('onAnnotationPopupAction')).toBe(1);

    await simulator.fireAnnotationEvent('onAnnotationInserted', { ids: ['a1'] });
    await simulator.fireAnnotationEvent('onAnnotationClicked', { id: 'a1' });
    await simulator.fireAnnotationEvent('onAnnotationPopupAction', {
      id: 'a1',
      action: 'Accept',
      critiqueSuggestion: '99.9%',
    });
    await simulator.fireAnnotationEvent('onAnnotationRemoved', { ids: ['a1'] });
    expect(events.filter((e) => e.type.startsWith('annotation-'))).toEqual([
      { type: 'annotation-inserted', surface: 'word', origin: 'local', annotationId: 'a1' },
      { type: 'annotation-clicked', surface: 'word', origin: 'local', annotationId: 'a1' },
      {
        type: 'annotation-action',
        surface: 'word',
        origin: 'local',
        annotationId: 'a1',
        action: 'Accept',
        suggestion: '99.9%',
      },
      { type: 'annotation-removed', surface: 'word', origin: 'local', annotationId: 'a1' },
    ]);

    unsub();
    await Promise.resolve();
    await Promise.resolve();
    expect(simulator.annotationHandlerCount('onAnnotationClicked')).toBe(0);
    expect(simulator.annotationHandlerCount('onAnnotationPopupAction')).toBe(0);

    // A 1.7 host registers the lifecycle events but not the 1.8 popup action; 1.6 registers none.
    simulator.restore();
    simulator = installFakeWord(seed(), { WordApi: 7 });
    new WordBridge().watch(() => undefined);
    await Promise.resolve();
    expect(simulator.annotationHandlerCount('onAnnotationInserted')).toBe(1);
    expect(simulator.annotationHandlerCount('onAnnotationPopupAction')).toBe(0);
    simulator.restore();
    simulator = installFakeWord(seed(), { WordApi: 6 });
    new WordBridge().watch(() => undefined);
    await Promise.resolve();
    expect(simulator.annotationHandlerCount('onAnnotationInserted')).toBe(0);
  });

  it('highlights the revealed anchor on 1.8 without touching content, and clearPreview removes it', async () => {
    simulator = installFakeWord(seed(), { WordApi: 8 });
    const bridge = new WordBridge();
    const before = simulator.snapshot().bodyText;
    await bridge.revealContext(clauseRef);
    expect(simulator.snapshot().highlights).toEqual(['99.5%']);
    expect(simulator.snapshot().bodyText).toBe(before);
    await bridge.clearPreview();
    expect(simulator.snapshot().highlights).toEqual([]);
    // Below 1.8 the reveal still happens, but no highlight is attempted.
    simulator.restore();
    simulator = installFakeWord(seed(), { WordApi: 7 });
    await new WordBridge().revealContext(clauseRef);
    expect(simulator.snapshot().highlights).toEqual([]);
  });
});
