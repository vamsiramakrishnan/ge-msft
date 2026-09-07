/** Metadata-only benchmark aggregation. Nearest-rank percentiles include only finite durations. */
export interface DurationSummary {
  samples: number;
  p50Ms?: number;
  p95Ms?: number;
}

export function summarizeDurations(values: ReadonlyArray<number | undefined>): DurationSummary {
  const sorted = values
    .filter((value): value is number => value !== undefined && Number.isFinite(value) && value >= 0)
    .sort((a, b) => a - b);
  return {
    samples: sorted.length,
    ...(sorted.length
      ? {
          p50Ms: sorted[Math.ceil(sorted.length * 0.5) - 1],
          p95Ms: sorted[Math.ceil(sorted.length * 0.95) - 1],
        }
      : {}),
  };
}

/** Each stage has its own denominator: a missing first token is not a zero-latency success. */
export function summarizeStages<T extends object>(
  measurements: readonly T[],
  stages: readonly (keyof T)[],
): Record<string, DurationSummary> {
  return Object.fromEntries(
    stages.map((stage) => [
      stage,
      summarizeDurations(
        measurements.map((measurement) => {
          const value = measurement[stage];
          return typeof value === 'number' ? value : undefined;
        }),
      ),
    ]),
  );
}
