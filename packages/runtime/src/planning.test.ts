import { describe, it, expect } from 'vitest';
import type { ActuationRequest, ChangeId } from '@ge/contracts';
import { approvalClassOf, isReversibleKind } from '@ge/contracts';
import { analyseEffectDependencies, propagateFailure, effectResources } from './planning.js';

const cid = (s: string) => s as ChangeId;

function req(
  kind: ActuationRequest['kind'],
  params: ActuationRequest['params'],
  id: string,
): ActuationRequest {
  return { changeId: cid(id), kind, surface: 'excel', params };
}

describe('planning — dependency DAG (ADR-0008 §7)', () => {
  it('infers table & chart depend on the spill (derived range), not on each other', () => {
    // spill a 10×2 table at Report!A1 → occupies A1:B11; table & chart read that range.
    const cells = Array.from({ length: 10 }, () => ['a', 'b']);
    const plan = analyseEffectDependencies([
      req('write-cells', { target: { range: 'Report!A1' }, cells }, 'c1'),
      req('create-table', { table: { range: 'Report!A1:B11', hasHeaders: true } }, 'c2'),
      req(
        'insert-chart',
        { chart: { chartType: 'column', sourceRange: 'Report!A1:B11', seriesBy: 'auto' } },
        'c3',
      ),
    ]);
    expect(plan.map((n) => n.id)).toEqual(['e1', 'e2', 'e3']);
    expect(plan[0]!.dependsOn).toEqual([]); // spill depends on nothing
    expect(plan[1]!.dependsOn).toEqual(['e1']); // table ← spill
    expect(plan[2]!.dependsOn).toEqual(['e1']); // chart ← spill (NOT chart ← table)
  });

  it('a non-overlapping effect is independent', () => {
    const plan = analyseEffectDependencies([
      req('write-cells', { target: { range: 'Report!A1' }, cells: [['x']] }, 'c1'),
      req('format-cells', { target: { range: 'Other!Z9' }, format: { bold: true } }, 'c2'),
    ]);
    expect(plan[1]!.dependsOn).toEqual([]);
  });

  it('propagateFailure skips the transitive dependents only', () => {
    const cells = Array.from({ length: 10 }, () => ['a', 'b']);
    const plan = analyseEffectDependencies([
      req('write-cells', { target: { range: 'Report!A1' }, cells }, 'c1'),
      req('create-table', { table: { range: 'Report!A1:B11', hasHeaders: true } }, 'c2'),
      req(
        'insert-chart',
        { chart: { chartType: 'column', sourceRange: 'Report!A1:B11', seriesBy: 'auto' } },
        'c3',
      ),
      req('format-cells', { target: { range: 'Other!Z9' }, format: { bold: true } }, 'c4'),
    ]);
    const { skipped, reason } = propagateFailure(plan, 'e1');
    expect(reason).toBe('prerequisite_failed');
    expect(skipped.sort()).toEqual(['e2', 'e3']); // table+chart skipped; the independent format (e4) is not
  });

  it('classifies approval authority + reversibility per kind', () => {
    const plan = analyseEffectDependencies([
      req('write-cells', { target: { range: 'A1' }, cells: [['x']] }, 'c1'),
      req('post-message', { text: 'hi' }, 'c2'),
      req('create-event', { appointment: { subject: 'sync' } }, 'c3'),
      req('resolve-revisions', { revisions: { scope: 'all', action: 'accept' } }, 'c4'),
    ]);
    expect(plan[0]!.approvalClass).toBe('in-document');
    expect(plan[1]!.approvalClass).toBe('external');
    expect(plan[2]!.approvalClass).toBe('estate');
    expect(plan[3]!.approvalClass).toBe('irreversible');
    expect(plan[3]!.reversible).toBe(false);
    expect(plan[0]!.reversible).toBe(true);
  });

  it('treats direct Word insertion kinds as irreversible until they return durable inverses', () => {
    expect(isReversibleKind('insert-text')).toBe(false);
    expect(isReversibleKind('insert-ooxml')).toBe(false);
    expect(approvalClassOf('insert-text')).toBe('irreversible');
    expect(approvalClassOf('insert-ooxml')).toBe('irreversible');
  });

  it('carries the changeId as the idempotency key', () => {
    const [node] = analyseEffectDependencies([
      req('write-cells', { target: { range: 'A1' }, cells: [['x']] }, 'idem-1'),
    ]);
    expect(node!.idempotencyKey).toBe('idem-1');
  });

  it('extracts the spilled region as the write resource (origin expanded by the grid)', () => {
    const cells = Array.from({ length: 3 }, () => ['a', 'b', 'c', 'd']); // 3×4
    const { writes } = effectResources(
      req('write-cells', { target: { range: 'Report!B2' }, cells }, 'c1'),
    );
    expect(writes).toEqual([{ kind: 'range', id: 'Report!B2:E4' }]);
  });
});
