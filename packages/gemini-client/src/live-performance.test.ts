import { describe, expect, it } from 'vitest';
import { summarizeDurations, summarizeStages } from './live-performance.js';

describe('live benchmark duration summaries', () => {
  it('does not turn failed or missing stages into zero-latency samples', () => {
    expect(summarizeDurations([undefined, NaN, Infinity, -1])).toEqual({ samples: 0 });
    expect(summarizeDurations([40, 0, 20, undefined, 30, 10])).toEqual({
      samples: 5,
      p50Ms: 20,
      p95Ms: 40,
    });
  });

  it('uses nearest-rank p95 for small and large samples without mutating evidence', () => {
    const values = Array.from({ length: 100 }, (_, index) => 100 - index);
    expect(summarizeDurations(values)).toEqual({ samples: 100, p50Ms: 50, p95Ms: 95 });
    expect(values[0]).toBe(100);
    expect(summarizeDurations([12])).toEqual({ samples: 1, p50Ms: 12, p95Ms: 12 });
  });

  it('separates stage denominators and keeps approval wait out of compute duration', () => {
    expect(
      summarizeStages(
        [
          { computeMs: 2, approvalWaitMs: 90, firstTokenMs: 15 },
          { computeMs: 3, approvalWaitMs: 500 },
        ],
        ['computeMs', 'approvalWaitMs', 'firstTokenMs'],
      ),
    ).toEqual({
      computeMs: { samples: 2, p50Ms: 2, p95Ms: 3 },
      approvalWaitMs: { samples: 2, p50Ms: 90, p95Ms: 500 },
      firstTokenMs: { samples: 1, p50Ms: 15, p95Ms: 15 },
    });
  });
});
