import { describe, it, expect } from 'vitest';
import {
  GroundingSelectionSchema,
  GROUND_SOURCE_TO_SELECTION_KIND,
  GroundSourceSchema,
  type GroundingSelection,
} from './index.js';

describe('GroundingSelection — the typed @-mention value', () => {
  it('accepts each reference kind (no id) and each addressable kind (with id)', () => {
    const ok: GroundingSelection[] = [
      { kind: 'current-context' },
      { kind: 'unit' },
      { kind: 'document', id: 'projects/x/dataStores/d/documents/42' },
      { kind: 'person', id: 'vamsi@acme' },
      { kind: 'data-store', id: 'projects/x/dataStores/d' },
      { kind: 'upload', fileId: 'file_abc' },
    ];
    for (const sel of ok) {
      expect(() => GroundingSelectionSchema.parse(sel)).not.toThrow();
    }
  });

  it('rejects an addressable kind with a missing or empty id', () => {
    expect(() => GroundingSelectionSchema.parse({ kind: 'document' })).toThrow();
    expect(() => GroundingSelectionSchema.parse({ kind: 'document', id: '' })).toThrow();
    expect(() => GroundingSelectionSchema.parse({ kind: 'data-store', id: '' })).toThrow();
    expect(() => GroundingSelectionSchema.parse({ kind: 'upload', fileId: '' })).toThrow();
  });

  it('rejects an unknown kind', () => {
    expect(() => GroundingSelectionSchema.parse({ kind: 'web' })).toThrow();
  });

  it('does not let id leak onto the reference kinds (strict discriminant)', () => {
    // current-context / unit carry no addressable id — the bridge attaches them.
    const parsed = GroundingSelectionSchema.parse({ kind: 'current-context' });
    expect(parsed).toEqual({ kind: 'current-context' });
  });

  it('maps every GroundSource kind onto a selection kind (total)', () => {
    for (const source of GroundSourceSchema.options) {
      expect(GROUND_SOURCE_TO_SELECTION_KIND[source]).toBeDefined();
    }
    expect(GROUND_SOURCE_TO_SELECTION_KIND.this).toBe('current-context');
    expect(GROUND_SOURCE_TO_SELECTION_KIND.datastore).toBe('data-store');
    expect(GROUND_SOURCE_TO_SELECTION_KIND.unit).toBe('unit');
  });
});
