import { describe, expect, it } from 'vitest';
import { CellGridSchema, CellSnapshotSchema } from './analysis.js';
import { ActuationRequestSchema } from './capability.js';

describe('cell grid boundary validation', () => {
  it.each([[], [[]], [[1], [2, 3]], null])(
    'rejects malformed grids without throwing: %j',
    (values) => {
      expect(CellGridSchema.safeParse(values).success).toBe(false);
      expect(
        CellSnapshotSchema.safeParse({
          surface: 'excel',
          documentId: 'doc',
          locator: 'Sheet1!A1',
          hash: `sha256:${'a'.repeat(64)}`,
          capturedAt: '2026-09-07T00:00:00.000Z',
          values,
          formulas: [],
        }).success,
      ).toBe(false);
      expect(
        ActuationRequestSchema.safeParse({
          kind: 'write-cells',
          surface: 'excel',
          changeId: 'invalid-grid',
          params: { target: { range: 'Sheet1!A1' }, cellValues: values, cellFormulas: [] },
        }).success,
      ).toBe(false);
    },
  );
});
