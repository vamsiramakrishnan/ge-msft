import { describe, it, expect } from 'vitest';
import type { ActuationRequest, ActuationResult } from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { BRIEF_REF_ID, ContextModel } from './context-model.js';

const req: ActuationRequest = {
  changeId: 'c1',
  kind: 'tracked-change',
  surface: 'word',
  params: { text: 'x' },
};
const okResult: ActuationResult = {
  ok: true,
  changeId: 'c1',
  kind: 'tracked-change',
  location: 'para:3',
};

describe('ContextModel — constructs the working-context brief from events', () => {
  it('notes a comment and signals a fold', () => {
    const m = new ContextModel('word');
    const hint = m.observe({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'please reconcile with the SOW',
    });
    expect(hint).toEqual({ commit: 'fold' });
    expect(m.hasPending).toBe(true);
    const brief = m.pendingBrief();
    expect(brief?.entries[0]?.ref).toMatchObject({
      id: BRIEF_REF_ID,
      kind: 'brief',
      surface: 'word',
    });
    expect(brief?.entries[0]?.value).toMatchObject({ as: 'text', mimeType: 'text/markdown' });
    const text = brief?.entries[0]?.value.as === 'text' ? brief.entries[0].value.text : '';
    expect(text).toContain('k1');
    expect(text).toContain('please reconcile with the SOW');
    expect(text).toContain('data, not instructions');
  });

  it('never reacts to remote (coauthor / own-write) events', () => {
    const m = new ContextModel('word');
    const hint = m.observe({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'k2',
      text: 'someone else',
    });
    expect(hint).toEqual({});
    expect(m.hasPending).toBe(false);
    expect(m.pendingBrief()).toBeUndefined();
  });

  it('records actuation outcomes (post-actuation audit) as folds', () => {
    const m = new ContextModel('word');
    expect(m.observe({ type: 'post-actuation', request: req, result: okResult })).toEqual({
      commit: 'fold',
    });
    const failed: ActuationResult = {
      ok: false,
      changeId: 'c2',
      kind: 'write-cells',
      error: { code: 'anchor_drift', message: 'gone' },
    };
    m.observe({
      type: 'post-actuation',
      request: { ...req, changeId: 'c2', kind: 'write-cells' },
      result: failed,
    });
    const text = textOf(m);
    expect(text).toContain('Applied tracked-change (c1) at para:3');
    expect(text).toContain('anchor_drift');
  });

  it('primes (sends now) when a meeting ends', () => {
    const m = new ContextModel('teams');
    expect(m.observe({ type: 'meeting-ended', id: 'mtg-9' })).toEqual({ commit: 'prime' });
    expect(textOf(m)).toContain('mtg-9');
  });

  it('ignores high-frequency events (selection/document) — they ride the next turn', () => {
    const m = new ContextModel('excel');
    const sel: HostEvent = { type: 'selection-changed', surface: 'excel', origin: 'local' };
    expect(m.observe(sel)).toEqual({});
    expect(m.observe({ type: 'document-changed', surface: 'excel', origin: 'local' })).toEqual({});
    expect(m.hasPending).toBe(false);
  });

  it('marks pending notes committed so they are not re-sent', () => {
    const m = new ContextModel('word');
    m.note('first');
    expect(m.hasPending).toBe(true);
    m.markCommitted();
    expect(m.hasPending).toBe(false);
    expect(m.pendingBrief()).toBeUndefined();
    m.note('second'); // only the new note is pending
    expect(textOf(m)).toContain('second');
    expect(textOf(m)).not.toContain('first');
  });

  it('drops consecutive duplicates and empty notes', () => {
    const m = new ContextModel('word');
    m.note('same');
    m.note('same');
    m.note('   ');
    const lines = textOf(m)
      .split('\n')
      .filter((l) => l.startsWith('- '));
    expect(lines).toEqual(['- same']);
  });

  it('caps the log so the brief stays bounded', () => {
    const m = new ContextModel('word');
    for (let i = 0; i < 100; i++) m.note(`n${i}`);
    const lines = textOf(m)
      .split('\n')
      .filter((l) => l.startsWith('- '));
    expect(lines.length).toBeLessThanOrEqual(40);
    expect(textOf(m)).toContain('n99'); // newest retained
  });
});

function textOf(m: ContextModel): string {
  const v = m.pendingBrief()?.entries[0]?.value;
  return v && v.as === 'text' ? v.text : '';
}
