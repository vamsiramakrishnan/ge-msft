import { describe, expect, it, vi } from 'vitest';
import { settleCalculation, suspendCalculation } from './calculation.js';

describe('calculation verification', () => {
  it('waits for Done without forcing a recalculation in Automatic mode', async () => {
    const application = {
      calculationState: 'Pending',
      calculationMode: 'Automatic',
      load: vi.fn(),
      calculate: vi.fn(),
    };
    let syncs = 0;
    await settleCalculation(application, async () => {
      // sync 1 loads calculationMode; syncs 2-3 poll calculationState.
      if (++syncs === 3) application.calculationState = 'Done';
    });
    expect(syncs).toBe(3);
    expect(application.load).toHaveBeenCalledWith('calculationMode');
    expect(application.calculate).not.toHaveBeenCalled();
  });
  it('forces Recalculate first when calculation mode is not Automatic', async () => {
    const application = {
      calculationState: 'Pending',
      calculationMode: 'Manual',
      load: vi.fn(),
      calculate: vi.fn(),
    };
    let syncs = 0;
    await settleCalculation(application, async () => {
      if (++syncs === 2) application.calculationState = 'Done';
    });
    expect(application.calculate).toHaveBeenCalledWith('Recalculate');
    expect(application.calculate).toHaveBeenCalledTimes(1);
  });
  it('treats an unknown calculation mode as non-automatic (fail safe: recalculate)', async () => {
    const application = { calculationState: 'Done', load: vi.fn(), calculate: vi.fn() };
    await settleCalculation(application, async () => undefined);
    expect(application.calculate).toHaveBeenCalledWith('Recalculate');
  });
  it('has a finite poll budget for a calculation that never settles', async () => {
    const application = {
      calculationState: 'Calculating',
      calculationMode: 'Automatic',
      load: vi.fn(),
      calculate: vi.fn(),
    };
    const sync = vi.fn(async () => undefined);
    await expect(settleCalculation(application, sync, 2)).rejects.toThrow('verification budget');
    // One sync for the mode load plus the two polls.
    expect(sync).toHaveBeenCalledTimes(3);
  });
});

describe('suspendCalculation', () => {
  const withOffice = (minor: number, run: () => void): void => {
    const g = globalThis as { Office?: unknown };
    const prev = g.Office;
    g.Office = {
      context: {
        requirements: {
          isSetSupported: (name: string, version?: string) =>
            name === 'ExcelApi' && parseFloat(version?.split('.')[1] ?? '0') <= minor,
        },
      },
    };
    try {
      run();
    } finally {
      g.Office = prev;
    }
  };
  it('suspends API calculation until the next sync on ExcelApi 1.6+', () => {
    const suspend = vi.fn();
    withOffice(6, () => {
      suspendCalculation({
        application: { suspendApiCalculationUntilNextSync: suspend },
      } as unknown as Excel.RequestContext);
    });
    expect(suspend).toHaveBeenCalledTimes(1);
  });
  it('is a no-op below ExcelApi 1.6 and tolerates a throwing host', () => {
    const suspend = vi.fn();
    withOffice(5, () => {
      suspendCalculation({
        application: { suspendApiCalculationUntilNextSync: suspend },
      } as unknown as Excel.RequestContext);
    });
    expect(suspend).not.toHaveBeenCalled();
    withOffice(9, () => {
      expect(() =>
        suspendCalculation({
          application: {
            suspendApiCalculationUntilNextSync: () => {
              throw new Error('refused');
            },
          },
        } as unknown as Excel.RequestContext),
      ).not.toThrow();
    });
  });
});
