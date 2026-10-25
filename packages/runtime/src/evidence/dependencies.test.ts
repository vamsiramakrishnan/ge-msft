import { describe, expect, it, vi } from 'vitest';
import type { DocumentPage } from '@ge/contracts';
import { EvidenceWorkspace } from './workspace.js';
import { MemoryEvidenceStore } from './store.js';
import { EvidenceDependencyIndex } from './dependencies.js';

function fixture(count = 10) {
  const pages: DocumentPage[] = Array.from({ length: count }, (_, index) => ({
    scope: 'native formula rows',
    limitations: [],
    units: [
      {
        id: `u:${index}`,
        text: String(index),
        relationsComplete: true,
        ...(index ? { relatedIds: [`u:${index - 1}`] } : {}),
      },
    ],
    ...(index + 1 < count ? { nextCursor: String(index + 1) } : {}),
  }));
  return {
    pages,
    source: {
      id: 'same-principal-document',
      readPage: async ({ cursor }: { cursor?: string }) =>
        structuredClone(pages[Number(cursor ?? 0)]!),
    },
  };
}

describe('bounded revision dependency indexing', () => {
  it('indexes eight pages per call and reuses unchanged authenticated pages between revisions', async () => {
    const { pages, source } = fixture();
    const store = new MemoryEvidenceStore();
    const workspace = new EvidenceWorkspace(store);
    const previous = await workspace.capture(source);
    const pageRead = vi.spyOn(store, 'getPage');
    const first = await workspace.indexDependencies(previous.id);
    expect(first).toMatchObject({ indexedPages: 8, complete: false, omittedUnits: 2 });
    expect(pageRead).toHaveBeenCalledTimes(8);
    expect(await workspace.indexDependencies(previous.id)).toMatchObject({
      complete: true,
      indexedUnits: 10,
    });
    pages[2]!.units[0]!.text = 'coauthor edit';
    const current = await workspace.capture(source);
    pageRead.mockClear();
    await workspace.indexDependencies(current.id);
    const final = await workspace.indexDependencies(current.id);
    expect(final).toMatchObject({ complete: true, reusedPages: 9 });
    expect(pageRead).toHaveBeenCalledTimes(1);
    const impact = workspace.dependencyImpact(previous.id, current.id);
    expect(impact.changedIds).toEqual(['u:2']);
    expect(impact.affectedIds).toEqual(['u:2', 'u:3', 'u:4', 'u:5', 'u:6', 'u:7', 'u:8', 'u:9']);
    expect(impact).toMatchObject({ complete: true, requiresFullRefresh: false });
    // The index is only a navigation accelerator. It does not relax full-source freshness.
    await expect(workspace.verify(previous.id, source)).rejects.toThrow('changed');
    expect(
      (
        await workspace.query(current.id, { kind: 'expand', id: 'u:2', direction: 'dependents' })
      ).nodes.map((node) => node.id),
    ).toEqual(['u:3']);
  });
  it('requires complete-source refresh when indexing, edge coverage or impact is partial', async () => {
    const { pages, source } = fixture(3);
    const workspace = new EvidenceWorkspace();
    const before = await workspace.capture(source);
    await workspace.indexDependencies(before.id, { maxPages: 1 });
    pages[0]!.units[0]!.text = 'changed';
    pages[1]!.units[0]!.relationsComplete = false;
    pages[2]!.units[0]!.relatedIds = ['missing-native-reference'];
    const after = await workspace.capture(source);
    const index = await workspace.indexDependencies(after.id);
    expect(index).toMatchObject({ unresolvedLinks: 1, incompleteRelationUnits: 1 });
    expect(workspace.dependencyImpact(before.id, after.id)).toMatchObject({
      complete: false,
      requiresFullRefresh: true,
    });
    await workspace.indexDependencies(before.id);
    expect(workspace.dependencyImpact(before.id, after.id, { maxNodes: 1 })).toMatchObject({
      complete: false,
      requiresFullRefresh: true,
    });
  });
  it('treats legacy unknown relation coverage as partial and preserves cycle safety', async () => {
    const { pages, source } = fixture(2);
    pages[0]!.units[0]!.relatedIds = ['u:1'];
    delete pages[0]!.units[0]!.relationsComplete;
    const workspace = new EvidenceWorkspace();
    const before = await workspace.capture(source);
    await workspace.indexDependencies(before.id);
    pages[1]!.units[0]!.text = 'new';
    const after = await workspace.capture(source);
    await workspace.indexDependencies(after.id);
    const result = workspace.dependencyImpact(before.id, after.id);
    expect(new Set(result.affectedIds)).toEqual(new Set(['u:1', 'u:0']));
    expect(result.requiresFullRefresh).toBe(true);
  });
  it('enforces aggregate index quotas without labelling a partial graph complete', async () => {
    const { source } = fixture(2);
    const workspace = new EvidenceWorkspace();
    const receipt = await workspace.capture(source);
    const index = new EvidenceDependencyIndex({
      maxRevisions: 1,
      maxUnits: 1,
      maxEdges: 100,
      maxBytes: 4096,
    });
    const progress = await index.advance(receipt, (page) => workspace.page(receipt.id, page));
    expect(progress).toMatchObject({
      complete: false,
      limited: true,
      indexedUnits: 1,
      omittedUnits: 1,
    });
    const aborted = new AbortController();
    aborted.abort();
    await expect(
      index.advance(receipt, (page) => workspace.page(receipt.id, page), {
        signal: aborted.signal,
      }),
    ).rejects.toThrow();
    await index.release(receipt.id);
    expect(await index.advance(receipt, (page) => workspace.page(receipt.id, page))).toMatchObject({
      limited: true,
    });
  });
  it('refuses corrupt pages and serializes competing incremental advances', async () => {
    const { source } = fixture(2);
    const workspace = new EvidenceWorkspace();
    const receipt = await workspace.capture(source);
    const index = new EvidenceDependencyIndex();
    await expect(
      index.advance(receipt, async () => ({ scope: 'forged', units: [], limitations: [] })),
    ).rejects.toThrow('integrity');
    const read = vi.fn((page: number) => workspace.page(receipt.id, page));
    const results = await Promise.all([
      index.advance(receipt, read, { maxPages: 1 }),
      index.advance(receipt, read, { maxPages: 1 }),
    ]);
    expect(results[0]!.indexedPages).toBe(1);
    expect(results[1]!.complete).toBe(true);
    expect(read).toHaveBeenCalledTimes(2);
    await workspace.release(receipt.id);
    await expect(workspace.indexDependencies(receipt.id)).rejects.toThrow('unavailable');
  });
});
