import { describe, expect, it, vi } from 'vitest';
import { settleCalculation } from './calculation.js';

describe('calculation verification', () => {
  it('waits for Done rather than hashing a pending result', async () => {
    const application = { calculationState: 'Pending', load: vi.fn(), calculate: vi.fn() };
    let syncs = 0;
    await settleCalculation(application, async () => {
      if (++syncs === 2) application.calculationState = 'Done';
    });
    expect(syncs).toBe(2);
    expect(application.calculate).toHaveBeenCalledWith('Recalculate');
  });
  it('has a finite poll budget for a calculation that never settles', async () => {
    const application = { calculationState: 'Calculating', load: vi.fn(), calculate: vi.fn() };
    const sync = vi.fn(async () => undefined);
    await expect(settleCalculation(application, sync, 2)).rejects.toThrow('verification budget');
    expect(sync).toHaveBeenCalledTimes(2);
  });
});
