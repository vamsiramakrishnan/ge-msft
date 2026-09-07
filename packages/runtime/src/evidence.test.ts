import { describe, expect, it, vi } from 'vitest';
import { EvidencePipeline } from './evidence.js';
import type { SearchRequest } from '@ge/gemini-client';
import { RuntimeHooks } from './hooks.js';
const store = 'projects/p/locations/global/collections/default_collection/dataStores/selected';
const context = { taskId: 'task', surface: 'word' as const };
function fixture(requiredSupport?: number) {
  const search = vi.fn(async (_request: SearchRequest) => ({
    facets: [],
    results: [
      {
        id: 'a',
        documentName: `${store}/branches/0/documents/a`,
        title: 'Report A',
        snippet: 'Revenue rose.',
        uri: 'javascript:bad()',
      },
      {
        id: 'b',
        documentName: `${store}/branches/0/documents/b`,
        title: 'Report B',
        snippet: 'Margin fell.',
      },
      { id: 'duplicate', documentName: `${store}/branches/0/documents/a`, snippet: 'duplicate' },
      {
        id: 'outside',
        documentName: 'projects/other/documents/a',
        snippet: 'Not in selected scope',
      },
    ],
  }));
  const rank = vi.fn(async () => [
    { id: 'e2', score: 1, content: 'Do not use returned rank content' },
    { id: 'invented', score: 2 },
  ]);
  const check = vi.fn(async () => ({
    supportScore: 0.9,
    claims: [{ claimText: 'Revenue rose.', score: 0.9 }],
    citedChunks: [],
  }));
  const hooks = new RuntimeHooks();
  const pipeline = new EvidencePipeline({
    search: { search },
    rank: { rank },
    grounding: { check },
    requiredSupport,
  });
  pipeline.install(hooks);
  return { hooks, pipeline, search, rank, check };
}
describe('evidence assembly hooks', () => {
  it('never performs an unscoped search', async () => {
    const f = fixture();
    await f.hooks.run('message:received', { mode: 'chat', text: 'question' }, context);
    expect(f.search).not.toHaveBeenCalled();
  });
  it('preserves explicit scope, deduplicates, reranks only original excerpts and sanitizes links', async () => {
    const f = fixture();
    const entries = await f.hooks.run(
      'message:received',
      {
        mode: 'chat',
        text: 'question',
        dataStoreSpecs: [{ dataStore: store, filter: 'team:finance' }],
      },
      context,
    );
    expect(f.search.mock.calls[0]?.[0]).toMatchObject({
      dataStoreSpecs: [{ dataStore: store, filter: 'team:finance' }],
    });
    expect(entries).toHaveLength(2);
    expect(JSON.stringify(entries[0])).toContain('Margin fell');
    expect(JSON.stringify(entries)).not.toContain('Do not use');
    expect(f.pipeline.state().sources.find((s) => s.title === 'Report A')?.uri).toBeUndefined();
    await f.hooks.run('model:response', { text: 'Revenue rose.', route: 'chat' }, context);
    expect(f.pipeline.state()).toMatchObject({ status: 'checked', score: 0.9, checkedClaims: 1 });
  });
  it('keeps command verification separate and clears task-scoped facts', async () => {
    const f = fixture();
    await f.hooks.run(
      'message:received',
      { mode: 'command', text: 'task', dataStoreSpecs: [{ dataStore: store }] },
      context,
    );
    await f.hooks.run(
      'model:response',
      { text: '```cmd\nset A1 2\n```', route: 'command' },
      context,
    );
    expect(f.check).not.toHaveBeenCalled();
    await f.hooks.run(
      'task:finished',
      {
        outcome: {
          taskId: 'task',
          surface: 'word',
          mode: 'command',
          status: 'completed',
          startedAt: new Date().toISOString(),
          effects: [],
          modelTurns: 0,
          toolCalls: 0,
        },
      },
      context,
    );
    await f.hooks.run('model:response', { text: 'Answer', route: 'chat' }, context);
    expect(f.check).not.toHaveBeenCalled();
  });
  it('fails closed when a required evidence service fails', async () => {
    const f = fixture(0.8);
    f.search.mockRejectedValueOnce(new Error('unavailable'));
    await expect(
      f.hooks.run(
        'message:received',
        { mode: 'chat', text: 'q', dataStoreSpecs: [{ dataStore: store }] },
        context,
      ),
    ).rejects.toThrow('unavailable');
  });
  it('reports optional ranking/check failures accurately', async () => {
    const f = fixture();
    f.rank.mockRejectedValueOnce(new Error('rank unavailable'));
    await f.hooks.run(
      'message:received',
      { mode: 'chat', text: 'q', dataStoreSpecs: [{ dataStore: store }] },
      context,
    );
    f.check.mockRejectedValueOnce(new Error('check unavailable'));
    await f.hooks.run('model:response', { text: 'answer', route: 'chat' }, context);
    expect(f.pipeline.state()).toMatchObject({ status: 'unavailable' });
  });
});
