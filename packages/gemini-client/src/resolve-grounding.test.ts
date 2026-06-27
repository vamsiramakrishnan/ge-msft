import { describe, it, expect } from 'vitest';
import type { GroundingSelection } from '@ge/contracts';
import { resolveGrounding, type GroundingResolveContext } from './resolve-grounding.js';
import type { QueryPart } from './session-context.js';

const selPart: QueryPart = { text: 'SLA is 99.5%' };
const unitDocPart: QueryPart = {
  documentReference: { documentName: 'projects/x/notebook-doc', displayTitle: 'Working doc' },
};

const ctx: GroundingResolveContext = {
  contextParts: [selPart],
  unitParts: [unitDocPart],
};

describe('resolveGrounding — typed @-mentions become request fields, never prompt text', () => {
  it('@this / current-context -> the context parts the bridge attaches', () => {
    const out = resolveGrounding([{ kind: 'current-context' }], ctx);
    expect(out.queryParts).toEqual([selPart]);
    expect(out.dataStoreSpecs).toBeUndefined();
    expect(out.fileIds).toBeUndefined();
  });

  it('unit -> the unit parts the bridge attaches', () => {
    const out = resolveGrounding([{ kind: 'unit' }], ctx);
    expect(out.queryParts).toEqual([unitDocPart]);
  });

  it('a data-store id -> toolsSpec dataStoreSpecs (not a query part)', () => {
    const out = resolveGrounding(
      [{ kind: 'data-store', id: 'projects/x/dataStores/dealroom' }],
      ctx,
    );
    expect(out.dataStoreSpecs).toEqual([{ dataStore: 'projects/x/dataStores/dealroom' }]);
    expect(out.queryParts).toBeUndefined();
    expect(out.fileIds).toBeUndefined();
  });

  it('an upload -> fileIds (not a query part, not a data store)', () => {
    const out = resolveGrounding([{ kind: 'upload', fileId: 'file_abc' }], ctx);
    expect(out.fileIds).toEqual(['file_abc']);
    expect(out.queryParts).toBeUndefined();
    expect(out.dataStoreSpecs).toBeUndefined();
  });

  it('a document id -> query.parts documentReference', () => {
    const out = resolveGrounding(
      [{ kind: 'document', id: 'projects/x/dataStores/d/branches/0/documents/42' }],
      ctx,
    );
    expect(out.queryParts).toEqual([
      {
        documentReference: { documentName: 'projects/x/dataStores/d/branches/0/documents/42' },
      },
    ]);
  });

  it('a person id -> query.parts personReference', () => {
    const out = resolveGrounding([{ kind: 'person', id: 'vamsi@acme' }], ctx);
    expect(out.queryParts).toEqual([{ personReference: { displayName: 'vamsi@acme' } }]);
  });

  it('combines kinds into the right buckets, order preserved within a bucket', () => {
    const selections: GroundingSelection[] = [
      { kind: 'current-context' },
      { kind: 'document', id: 'doc1' },
      { kind: 'data-store', id: 'ds1' },
      { kind: 'upload', fileId: 'f1' },
      { kind: 'person', id: 'p1' },
      { kind: 'data-store', id: 'ds2' },
    ];
    const out = resolveGrounding(selections, ctx);
    expect(out.queryParts).toEqual([
      selPart,
      { documentReference: { documentName: 'doc1' } },
      { personReference: { displayName: 'p1' } },
    ]);
    expect(out.dataStoreSpecs).toEqual([{ dataStore: 'ds1' }, { dataStore: 'ds2' }]);
    expect(out.fileIds).toEqual(['f1']);
  });

  it('is pure and total — never throws, no input mutation', () => {
    const selections: GroundingSelection[] = [{ kind: 'data-store', id: 'ds1' }];
    const frozen = Object.freeze([...selections]);
    expect(() => resolveGrounding(frozen, ctx)).not.toThrow();
    expect(selections).toEqual([{ kind: 'data-store', id: 'ds1' }]);
  });

  it('drops an unresolvable selection with a structured note, never throws or inlines text', () => {
    // current-context picked, but the bridge attached no context parts: unresolvable.
    const out = resolveGrounding([{ kind: 'current-context' }, { kind: 'unit' }], {});
    expect(out.queryParts).toBeUndefined();
    expect(out.dataStoreSpecs).toBeUndefined();
    expect(out.fileIds).toBeUndefined();
    expect(out.notes).toBeDefined();
    expect(out.notes!.map((n) => n.kind)).toEqual(['current-context', 'unit']);
    for (const note of out.notes!) {
      expect(note.reason).toBeTypeOf('string');
    }
  });

  it('NO grounding selection is silently turned into raw prompt text', () => {
    const selections: GroundingSelection[] = [
      { kind: 'current-context' },
      { kind: 'unit' },
      { kind: 'document', id: 'doc1' },
      { kind: 'person', id: 'p1' },
      { kind: 'data-store', id: 'ds1' },
      { kind: 'upload', fileId: 'f1' },
    ];
    const out = resolveGrounding(selections, ctx);
    // Every text-bearing query part must be a typed reference part the bridge attached
    // (current-context/unit) — never a free-text part synthesized from a selection.
    const rawTextParts = (out.queryParts ?? []).filter(
      (p) => 'text' in p && !ctx.contextParts!.includes(p) && !ctx.unitParts!.includes(p),
    );
    expect(rawTextParts).toEqual([]);
  });

  it('empty selections -> empty partial', () => {
    expect(resolveGrounding([], ctx)).toEqual({});
  });
});
