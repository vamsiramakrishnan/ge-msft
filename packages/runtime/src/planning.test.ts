import { describe, it, expect } from 'vitest';
import type { ActuationRequest, ChangeId } from '@ge/contracts';
import { approvalClassOf, isReversibleKind } from '@ge/contracts';
import {
  analyseEffectDependencies,
  propagateFailure,
  effectResources,
  changesCoveredBy,
} from './planning.js';

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

  it('a PowerPoint slide=last write depends on a preceding insert-slide; an addressed one does not', () => {
    const pp = (
      kind: ActuationRequest['kind'],
      params: ActuationRequest['params'],
      id: string,
    ) => ({
      ...req(kind, params, id),
      surface: 'powerpoint' as const,
    });
    const plan = analyseEffectDependencies([
      pp('insert-slide', { slide: { title: 'Q4 plan', bullets: [] } }, 'c1'),
      pp('add-shape', { target: { slideId: 'last' }, shape: { shapeType: 'textBox' } }, 'c2'),
      pp(
        'add-table-slide',
        { target: { slideId: 'slide:LAST' }, tableGrid: { hasHeaders: false, rows: [['a']] } },
        'c3',
      ),
    ]);
    expect(plan[1]!.dependsOn).toEqual(['e1']);
    expect(plan[2]!.dependsOn).toEqual(['e1']);
    // So a failed/uncertain insert skips the slide=last writes instead of running them.
    expect(propagateFailure(plan, 'e1').skipped).toEqual(['e2', 'e3']);

    // A write to an exact slide id does not depend on the insert.
    const addressed = analyseEffectDependencies([
      pp('insert-slide', { slide: { title: 'Q4 plan', bullets: [] } }, 'c1'),
      pp('add-shape', { target: { slideId: '256#0' }, shape: { shapeType: 'textBox' } }, 'c2'),
    ]);
    expect(addressed[1]!.dependsOn).toEqual([]);
  });

  it('ties every relative-slide PowerPoint write to the step that adds a slide', () => {
    const pp = (
      kind: ActuationRequest['kind'],
      params: ActuationRequest['params'],
      id: string,
    ) => ({
      ...req(kind, params, id),
      surface: 'powerpoint' as const,
    });
    // insert-slide, then shape writes addressed at the last slide (incl. pp:slide:last).
    const plan = analyseEffectDependencies([
      pp('insert-slide', { slide: { title: 'Risks', bullets: [] } }, 'c1'),
      pp('set-shape-text', { target: { slideId: 'last', shapeId: 'title' }, text: 'x' }, 'c2'),
      pp(
        'format-shape',
        { target: { slideId: 'last', shapeId: 'title' }, shapeFormat: { fill: '#000' } },
        'c3',
      ),
      pp(
        'add-shape',
        { target: { slideId: 'pp:slide:last' }, shape: { shapeType: 'textBox' } },
        'c4',
      ),
    ]);
    expect(plan.slice(1).map((node) => node.dependsOn)).toEqual([['e1'], ['e1'], ['e1']]);
    expect(propagateFailure(plan, 'e1').skipped).toEqual(['e2', 'e3', 'e4']);

    // `/add-table-slide slide=new` creates the last slide, so a later slide=last write depends on it.
    const table = analyseEffectDependencies([
      pp(
        'add-table-slide',
        { target: { slideId: 'new' }, tableGrid: { hasHeaders: false, rows: [['a']] } },
        'c1',
      ),
      pp('add-shape', { target: { slideId: 'last' }, shape: { shapeType: 'textBox' } }, 'c2'),
    ]);
    expect(table[1]!.dependsOn).toEqual(['e1']);
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

  it('makes a multi-area chart depend on a write into any of its areas', () => {
    // Security review L8: `Sheet2!C1:C11,Sheet2!G1:G11` did not parse, so the chart had no dependency.
    const plan = analyseEffectDependencies([
      req('write-cells', { target: { range: 'Sheet2!G2' }, cells: [['1']] }, 'c1'),
      req(
        'insert-chart',
        {
          chart: { chartType: 'bar', sourceRange: 'Sheet2!C1:C11,Sheet2!G1:G11', seriesBy: 'auto' },
        },
        'c2',
      ),
    ]);
    expect(plan[1]!.dependsOn).toEqual(['e1']);
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

describe('changesCoveredBy — cells an earlier effect of the batch produced', () => {
  const at = (locator: string, values: unknown[][], formulas?: unknown[][]) => ({
    locator,
    values,
    ...(formulas ? { formulas } : {}),
  });
  const before = at('S!A1:B2', [
    ['a', 1],
    ['b', 2],
  ]);
  const after = at('S!A1:B2', [
    ['a', 9],
    ['b', 2],
  ]);

  it('accepts a changed cell that holds what the batch wrote there', () => {
    expect(changesCoveredBy(before, after, [at('S!B1', [[9]])])).toBe(true);
    const quoted = { ...before, locator: "'My Sheet'!A1:B2" };
    expect(
      changesCoveredBy(quoted, { ...after, locator: quoted.locator }, [at('my sheet!B1', [[9]])]),
    ).toBe(true);
  });

  it('refuses a changed cell that holds something else than the batch wrote', () => {
    expect(changesCoveredBy(before, after, [at('S!B1', [[8]])])).toBe(false);
  });

  it('refuses a change outside every landed capture', () => {
    const edited = at('S!A1:B2', [
      ['a', 9],
      ['x', 2],
    ]);
    expect(changesCoveredBy(before, edited, [at('S!B1', [[9]])])).toBe(false);
  });

  it('refuses a capture on another sheet, or one without a sheet', () => {
    expect(changesCoveredBy(before, after, [at('T!B1', [[9]])])).toBe(false);
    expect(changesCoveredBy(before, after, [at('B1', [[9]])])).toBe(false);
    const obrien = { ...before, locator: "'O''Brien'!A1:B2" };
    expect(
      changesCoveredBy(obrien, { ...after, locator: obrien.locator }, [at('OBrien!B1', [[9]])]),
    ).toBe(false);
  });

  it('counts a formula change as a change', () => {
    const formula = at('S!A1:B2', before.values as unknown[][], [
      ['', '=1'],
      ['', ''],
    ]);
    expect(changesCoveredBy(before, formula, [])).toBe(false);
  });

  it('refuses a different shape, and captures with no changed cell', () => {
    expect(changesCoveredBy(before, at('S!A1:B2', [['a', 1]]), [at('S!A1:B2', [['a', 1]])])).toBe(
      false,
    );
    expect(changesCoveredBy(before, before, [at('S!B1', [[1]])])).toBe(false);
  });
});
