import { describe, it, expect, vi } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { ApprovalCoordinator } from './approval-coordinator.js';
import type { PendingWrite, PendingPlan } from './controller.js';

const write = (id: string): PendingWrite => ({
  changeId: asChangeId(id),
  kind: 'tracked-change',
  command: `rewrite ${id}`,
});

const plan = (summary: string): PendingPlan => ({ effects: [], summary });

/** A coordinator wired to spy callbacks so we can assert the card view-state pushes. */
const make = () => {
  const showWrite = vi.fn<[PendingWrite | undefined], void>();
  const showPlan = vi.fn<[PendingPlan | undefined], void>();
  const coord = new ApprovalCoordinator(showWrite, showPlan);
  return { coord, showWrite, showPlan };
};

describe('ApprovalCoordinator — write gate', () => {
  it('approves only on an explicit user approval', async () => {
    const { coord, showWrite } = make();
    const id = asChangeId('c-1');
    const decision = coord.awaitWrite(write('c-1'), id);
    expect(showWrite).toHaveBeenCalledWith(write('c-1'));
    coord.approveWrite(id);
    await expect(decision).resolves.toBe(true);
  });

  it('rejects and drops the card immediately', async () => {
    const { coord, showWrite } = make();
    const id = asChangeId('c-1');
    const decision = coord.awaitWrite(write('c-1'), id);
    coord.rejectWrite(id);
    await expect(decision).resolves.toBe(false);
    expect(showWrite).toHaveBeenLastCalledWith(undefined);
  });

  it('approve KEEPS the card until the write-result is consumed', async () => {
    const { coord, showWrite } = make();
    const id = asChangeId('c-1');
    const decision = coord.awaitWrite(write('c-1'), id);
    coord.approveWrite(id);
    await decision;
    // not cleared yet — the loop has not narrated the write-result
    expect(showWrite).toHaveBeenCalledTimes(1);
    coord.consumeWriteResult();
    expect(showWrite).toHaveBeenLastCalledWith(undefined);
  });

  it('consumeWriteResult fails closed if a decision is somehow still awaiting', async () => {
    const { coord, showWrite } = make();
    const decision = coord.awaitWrite(write('c-1'), asChangeId('c-1'));
    // out-of-order: consume the result while the decision is still open (should be impossible)
    coord.consumeWriteResult();
    await expect(decision).resolves.toBe(false); // settled fail-closed, not left dangling
    expect(showWrite).toHaveBeenLastCalledWith(undefined);
  });

  it('ignores a SUPERSEDED decision (late click on a replaced card)', async () => {
    const { coord } = make();
    const decision = coord.awaitWrite(write('c-2'), asChangeId('c-2'));
    coord.approveWrite(asChangeId('c-1')); // stale id — no-op
    let settled = false;
    void decision.then(() => (settled = true));
    await Promise.resolve();
    expect(settled).toBe(false);
    coord.approveWrite(asChangeId('c-2')); // current id resolves it
    await expect(decision).resolves.toBe(true);
  });

  it('treats an undefined id as a match (non-superseding approve)', async () => {
    const { coord } = make();
    const decision = coord.awaitWrite(write('c-1'), asChangeId('c-1'));
    coord.approveWrite();
    await expect(decision).resolves.toBe(true);
  });

  it('releaseAwaiting settles a pending write fail-closed', async () => {
    const { coord } = make();
    const decision = coord.awaitWrite(write('c-1'), asChangeId('c-1'));
    coord.releaseAwaiting();
    await expect(decision).resolves.toBe(false);
  });

  it('a fresh awaitWrite releases the prior still-open decision fail-closed', async () => {
    const { coord } = make();
    const first = coord.awaitWrite(write('c-1'), asChangeId('c-1'));
    const second = coord.awaitWrite(write('c-2'), asChangeId('c-2'));
    await expect(first).resolves.toBe(false);
    coord.approveWrite(asChangeId('c-2'));
    await expect(second).resolves.toBe(true);
  });
});

describe('ApprovalCoordinator — plan gate', () => {
  it('approves on an explicit plan approval', async () => {
    const { coord, showPlan } = make();
    const decision = coord.awaitPlan(plan('3 writes'));
    expect(showPlan).toHaveBeenCalledWith(plan('3 writes'));
    coord.approvePlan();
    await expect(decision).resolves.toBe(true);
    expect(showPlan).toHaveBeenLastCalledWith(undefined); // either decision drops the plan card
  });

  it('rejects the plan and drops the card', async () => {
    const { coord, showPlan } = make();
    const decision = coord.awaitPlan(plan('1 write'));
    coord.rejectPlan();
    await expect(decision).resolves.toBe(false);
    expect(showPlan).toHaveBeenLastCalledWith(undefined);
  });

  it('releaseAwaiting settles a pending plan fail-closed', async () => {
    const { coord } = make();
    const decision = coord.awaitPlan(plan('2 writes'));
    coord.releaseAwaiting();
    await expect(decision).resolves.toBe(false);
  });
});

describe('ApprovalCoordinator — releaseAll', () => {
  it('settles an awaiting write fail-closed and drops a lingering approved card', async () => {
    const { coord, showWrite } = make();
    const id = asChangeId('c-1');
    const decision = coord.awaitWrite(write('c-1'), id);
    coord.approveWrite(id); // approved but write-result not yet consumed → card lingers
    await decision;
    coord.releaseAll();
    expect(showWrite).toHaveBeenLastCalledWith(undefined);
  });

  it('settles BOTH gates fail-closed when called while awaiting', async () => {
    const { coord } = make();
    const w = coord.awaitWrite(write('c-1'), asChangeId('c-1'));
    const p = coord.awaitPlan(plan('1 write'));
    coord.releaseAll();
    await expect(w).resolves.toBe(false);
    await expect(p).resolves.toBe(false);
  });

  it('is a safe no-op when nothing is staged', () => {
    const { coord, showWrite, showPlan } = make();
    coord.releaseAll();
    expect(showWrite).not.toHaveBeenCalled();
    expect(showPlan).not.toHaveBeenCalled();
  });
});
