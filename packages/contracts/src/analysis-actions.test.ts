import { describe, expect, it } from 'vitest';
import {
  AnalysisActionSchema,
  ANALYSIS_BINDING_KINDS,
  isAnalysisBindingKind,
  ReconciliationSpecSchema,
} from './analysis-actions.js';
import { AnalysisProgramSchema } from './analysis-program-schema.js';

describe('shared analysis boundary schemas', () => {
  it('preserves input defaults and existing object stripping semantics', () => {
    expect(AnalysisActionSchema.parse({ kind: 'capture', range: 'A1:B2', ignored: true })).toEqual({
      kind: 'capture',
      range: 'A1:B2',
      headers: true,
    });
    expect(
      AnalysisActionSchema.parse({
        kind: 'query',
        sql: 'SELECT * FROM $source',
        inputs: ['$source'],
      }),
    ).toEqual({
      kind: 'query',
      sql: 'SELECT * FROM $source',
      inputs: ['$source'],
      title: 'Query result',
    });
    expect(
      ReconciliationSpecSchema.parse({
        left: '$left',
        right: '$right',
        leftKey: 0,
        rightKey: 0,
        leftAmount: 1,
        rightAmount: 1,
        currency: 'USD',
      }).tolerance,
    ).toBe('0.01');
  });

  it('preserves cross-field currency validation', () => {
    const spec = {
      left: '$left',
      right: '$right',
      leftKey: 0,
      rightKey: 0,
      leftAmount: 1,
      rightAmount: 1,
    };
    expect(ReconciliationSpecSchema.safeParse(spec).success).toBe(false);
    expect(ReconciliationSpecSchema.safeParse({ ...spec, leftCurrency: 2 }).success).toBe(false);
    expect(
      ReconciliationSpecSchema.safeParse({ ...spec, leftCurrency: 2, rightCurrency: 2 }).success,
    ).toBe(true);
    expect(ReconciliationSpecSchema.safeParse({ ...spec, currency: 'USD' }).success).toBe(true);
  });

  it('shares strict guard values and producer classification without granting effects binding authority', () => {
    expect(ANALYSIS_BINDING_KINDS).toEqual(['capture', 'query', 'reconcile', 'filter', 'inspect']);
    for (const kind of ANALYSIS_BINDING_KINDS) expect(isAnalysisBindingKind(kind)).toBe(true);
    for (const kind of ['materialize', 'remove', 'undo', 'resume', 'forget', 'recovery', undefined])
      expect(isAnalysisBindingKind(kind)).toBe(false);
    expect(
      AnalysisActionSchema.safeParse({
        kind: 'materialize',
        id: '$source',
        destination: 'B1',
        whenNonEmpty: 'true',
      }).success,
    ).toBe(false);
    expect(
      AnalysisActionSchema.safeParse({
        kind: 'query',
        inputs: ['$source'],
        sql: 'SELECT c0 FROM $source',
        requiredColumns: [{ input: '$source', indices: [0], exactDecimal: true, ignored: true }],
      }).success,
    ).toBe(false);
  });

  it('preserves bounded program shape and verified completion default', () => {
    const step = { op: 'bind', name: 'source', action: { kind: 'capture', range: 'A1:B2' } };
    expect(AnalysisProgramSchema.parse({ version: 1, steps: [step] })).toMatchObject({
      completion: 'verified',
      steps: [{ action: { headers: true } }],
    });
    expect(
      AnalysisProgramSchema.safeParse({ version: 1, steps: Array(32).fill(step) }).success,
    ).toBe(false);
    expect(
      AnalysisProgramSchema.safeParse({ version: 1, steps: [step], ignored: true }).success,
    ).toBe(false);
    expect(AnalysisProgramSchema.safeParse({ version: 2, steps: [step] }).success).toBe(false);
  });
});
