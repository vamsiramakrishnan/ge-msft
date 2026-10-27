import { describe, expect, it, vi } from 'vitest';
import { asChangeId, contentHash, type ActuationRequest } from '@ge/contracts';
import { MemoryTaskJournalStore, TaskJournal, type TaskJournalStore } from './task-journal.js';

const scope = 'sha256:' + 'a'.repeat(64);
const request = (id: string): ActuationRequest => ({
  changeId: asChangeId(id),
  surface: 'word',
  kind: 'tracked-change',
  params: { target: { matchText: 'before' }, text: 'after' },
});
describe('durable task checkpoints', () => {
  it('journals large binary operations without duplicating payloads or inverse bodies', async () => {
    const store = new MemoryTaskJournalStore();
    const journal = new TaskJournal(store, scope);
    const base64 = 'A'.repeat(3_000_000);
    const input: ActuationRequest = {
      changeId: asChangeId('deck'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: { deck: { format: 'pptx', base64 } },
    };
    const receipt = await journal.execute('task', input, async () => ({
      ok: true,
      changeId: input.changeId,
      kind: input.kind,
      inverse: { op: 'restore-slide', base64 },
    }));
    expect(receipt.inverse).toMatchObject({ base64 });
    const stored = JSON.stringify(await store.read(scope));
    expect(stored.length).toBeLessThan(2000);
    expect(stored).toContain('payloadOmitted');
    expect(stored).not.toContain('base64');
  });
  it('blocks a second pane while an earlier effect is in flight and preserves a thrown outcome', async () => {
    const store = new MemoryTaskJournalStore();
    const first = new TaskJournal(store, scope);
    const second = new TaskJournal(store, scope);
    let finish!: () => void;
    const barrier = new Promise<void>((resolve) => {
      finish = resolve;
    });
    const dispatch = vi.fn(async () => {
      await barrier;
      throw new Error('after dispatch');
    });
    const running = first.execute('task-1', request('first'), dispatch);
    await vi.waitFor(() => expect(dispatch).toHaveBeenCalledOnce());
    const duplicate = vi.fn();
    await expect(second.execute('task-2', request('second'), duplicate)).rejects.toThrow(
      'unresolved',
    );
    expect(duplicate).not.toHaveBeenCalled();
    finish();
    expect((await running).recoveryPending).toBe(true);
    const reopened = new TaskJournal(store, scope);
    expect(await reopened.inspect()).toMatchObject([{ state: 'uncertain', id: 'first' }]);
    await expect(reopened.execute('task-3', request('third'), duplicate)).rejects.toThrow(
      'unresolved',
    );
  });
  it('isolates principals/documents and never replays an already journaled change ID', async () => {
    const store = new MemoryTaskJournalStore();
    const journal = new TaskJournal(store, scope);
    const receipt = { ok: true, changeId: asChangeId('first'), kind: 'tracked-change' as const };
    await journal.execute('task-1', request('first'), async () => receipt);
    await expect(journal.execute('task-1', request('first'), vi.fn())).rejects.toThrow(
      'already journaled',
    );
    const other = new TaskJournal(
      store,
      await contentHash({ principal: 'other', document: 'other' }),
    );
    expect(await other.inspect()).toEqual([]);
  });
  it('does not dispatch when intent persistence fails', async () => {
    const store: TaskJournalStore = {
      durable: true,
      read: async () => [],
      update: async () => {
        throw new Error('storage full');
      },
    };
    const dispatch = vi.fn();
    await expect(
      new TaskJournal(store, scope).execute('task', request('first'), dispatch),
    ).rejects.toThrow('storage full');
    expect(dispatch).not.toHaveBeenCalled();
  });
  it('keeps the pending claim if the post-write checkpoint fails', async () => {
    const memory = new MemoryTaskJournalStore();
    let updates = 0;
    const store: TaskJournalStore = {
      durable: true,
      read: (key) => memory.read(key),
      update: (key, change) => {
        if (++updates === 2) throw new Error('disk');
        return memory.update(key, change);
      },
    };
    const journal = new TaskJournal(store, scope);
    const receipt = await journal.execute('task', request('first'), async () => ({
      ok: true,
      changeId: asChangeId('first'),
      kind: 'tracked-change',
    }));
    expect(receipt).toMatchObject({ ok: true, recoveryPending: true });
    expect(await journal.inspect()).toMatchObject([{ state: 'dispatching' }]);
    await expect(journal.execute('task', request('another'), vi.fn())).rejects.toThrow(
      'unresolved',
    );
  });
});

describe('stable idempotency keys (B4)', () => {
  const verified = (id: string) => ({
    ok: true,
    changeId: asChangeId(id),
    kind: 'tracked-change' as const,
    verification: { status: 'verified' as const },
  });
  it('finds the prior outcome of the same intent under a new changeId', async () => {
    const journal = new TaskJournal(new MemoryTaskJournalStore(), scope);
    await journal.execute('task', request('first'), async () => verified('first'));
    const prior = journal.priorOutcome(request('second'));
    expect(prior).toMatchObject({ id: 'first', state: 'applied' });
    expect(journal.priorOutcome({ ...request('x'), params: { text: 'other' } })).toBeUndefined();
  });
  it('replays a verified prior outcome only when asked, naming the original change', async () => {
    const journal = new TaskJournal(new MemoryTaskJournalStore(), scope);
    await journal.execute('task', request('first'), async () => verified('first'));
    const dispatch = vi.fn(async () => verified('second'));
    // The plain path dispatches again: a re-emitted write is a new request.
    const plain = await journal.execute('task', request('second'), dispatch);
    expect(dispatch).toHaveBeenCalledOnce();
    expect(plain.replayOf).toBeUndefined();
    // A resume opts in and gets the prior receipt back without a dispatch.
    const again = vi.fn(async () => verified('third'));
    const replayed = await journal.execute('task', request('third'), again, {
      replayIntent: true,
    });
    expect(again).not.toHaveBeenCalled();
    expect(replayed).toMatchObject({ ok: true, changeId: 'third', replayOf: 'second' });
    expect(journal.list().find((entry) => entry.id === 'third')).toMatchObject({
      state: 'applied',
      replayOf: 'second',
    });
  });
  it('never replays an unverified, uncertain or other-task outcome', async () => {
    const journal = new TaskJournal(new MemoryTaskJournalStore(), scope);
    await journal.execute('task', request('plain'), async () => ({
      ok: true,
      changeId: asChangeId('plain'),
      kind: 'tracked-change' as const,
    }));
    const dispatch = vi.fn(async () => verified('retry'));
    await journal.execute('task', request('retry'), dispatch, { replayIntent: true });
    expect(dispatch).toHaveBeenCalledOnce(); // prior was applied but not verified
    const other = vi.fn(async () => verified('other-task'));
    await journal.execute('task-2', request('other-task'), other, { replayIntent: true });
    expect(other).toHaveBeenCalledOnce(); // verified prior belongs to another task
  });
});
