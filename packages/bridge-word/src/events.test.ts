import { describe, it, expect } from 'vitest';
import {
  originFromWordSource,
  selectionChangedEvent,
  documentChangedEvent,
  commentAddedEvent,
} from './events.js';

describe('word event origin (pure)', () => {
  it('derives remote only when the host says remote', () => {
    expect(originFromWordSource('Remote')).toBe('remote');
    expect(originFromWordSource('remote')).toBe('remote');
    expect(originFromWordSource('REMOTE')).toBe('remote');
  });

  it('defaults to local for local source', () => {
    expect(originFromWordSource('Local')).toBe('local');
    expect(originFromWordSource('local')).toBe('local');
  });

  it('defaults to local for unknown / missing / non-string source', () => {
    expect(originFromWordSource(undefined)).toBe('local');
    expect(originFromWordSource(null)).toBe('local');
    expect(originFromWordSource('')).toBe('local');
    expect(originFromWordSource('something-else')).toBe('local');
    expect(originFromWordSource(42)).toBe('local');
    expect(originFromWordSource({})).toBe('local');
  });
});

describe('word event builders (pure)', () => {
  it('builds a selection-changed event (always local) with optional preview', () => {
    expect(selectionChangedEvent()).toEqual({
      type: 'selection-changed',
      surface: 'word',
      origin: 'local',
    });
    expect(selectionChangedEvent('hello world')).toEqual({
      type: 'selection-changed',
      surface: 'word',
      origin: 'local',
      preview: 'hello world',
    });
  });

  it('attaches an optional context ref to selection-changed', () => {
    const ref = {
      id: 'word:selection',
      kind: 'selection' as const,
      surface: 'word' as const,
      title: 'Selection',
    };
    const evt = selectionChangedEvent('hi', ref);
    expect(evt).toMatchObject({ type: 'selection-changed', surface: 'word', ref });
  });

  it('builds a document-changed event preserving the derived origin', () => {
    expect(documentChangedEvent('local')).toEqual({
      type: 'document-changed',
      surface: 'word',
      origin: 'local',
    });
    expect(documentChangedEvent('remote')).toEqual({
      type: 'document-changed',
      surface: 'word',
      origin: 'remote',
    });
  });

  it('builds a comment-added event with id and optional text', () => {
    expect(commentAddedEvent('local', 'cmt-1')).toEqual({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'cmt-1',
    });
    expect(commentAddedEvent('remote', 'cmt-2', 'please revise')).toEqual({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'cmt-2',
      text: 'please revise',
    });
  });
});
