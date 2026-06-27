import { describe, it, expect } from 'vitest';
import { selectionChanged, documentChanged } from './events.js';

describe('powerpoint events (pure)', () => {
  it('builds a local selection-changed event, carrying a preview when present', () => {
    expect(selectionChanged()).toEqual({
      type: 'selection-changed',
      surface: 'powerpoint',
      origin: 'local',
    });
    expect(selectionChanged('Slide 4')).toEqual({
      type: 'selection-changed',
      surface: 'powerpoint',
      origin: 'local',
      preview: 'Slide 4',
    });
  });

  it('omits an empty preview', () => {
    expect(selectionChanged('')).not.toHaveProperty('preview');
  });

  it('builds a document-changed event defaulting to local origin', () => {
    expect(documentChanged()).toEqual({
      type: 'document-changed',
      surface: 'powerpoint',
      origin: 'local',
    });
    expect(documentChanged('remote').origin).toBe('remote');
  });
});
