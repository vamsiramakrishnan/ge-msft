import { describe, it, expect } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import { planFormatCells, planAddComment } from './actuate-plan.js';

function req(params: ActuationRequest['params'], kind: ActuationRequest['kind']): ActuationRequest {
  return { changeId: asChangeId('c1'), kind, surface: 'excel', params };
}

describe('planFormatCells (ADR-0004 format-cells)', () => {
  it('maps each present format facet to a host op and flags hasOps', () => {
    const plan = planFormatCells(
      req(
        {
          target: { range: 'Sales!A1:C1' },
          format: { bold: true, italic: false, fill: '#FFF2CC', numberFormat: '$#,##0.00' },
        },
        'format-cells',
      ),
    );
    expect(plan).toEqual({
      address: 'Sales!A1:C1',
      bold: true,
      italic: false,
      fill: '#FFF2CC',
      numberFormat: '$#,##0.00',
      hasOps: true,
    });
  });

  it('omits absent facets (undefined ⇒ leave untouched) and still reports hasOps for one facet', () => {
    const plan = planFormatCells(
      req({ target: { range: 'A1' }, format: { bold: true } }, 'format-cells'),
    );
    expect(plan).toEqual({ address: 'A1', bold: true, hasOps: true });
    expect(plan.italic).toBeUndefined();
    expect(plan.fill).toBeUndefined();
    expect(plan.numberFormat).toBeUndefined();
  });

  it('reports hasOps:false for an empty/absent format (a no-op the bridge degrades)', () => {
    expect(planFormatCells(req({ target: { range: 'A1' } }, 'format-cells')).hasOps).toBe(false);
    expect(
      planFormatCells(req({ target: { range: 'A1' }, format: {} }, 'format-cells')).hasOps,
    ).toBe(false);
  });

  it('omits the address when no target.range is given', () => {
    const plan = planFormatCells(req({ format: { bold: true } }, 'format-cells'));
    expect(plan.address).toBeUndefined();
    expect(plan.hasOps).toBe(true);
  });

  it('preserves an explicit false (e.g. unbold) rather than dropping it', () => {
    const plan = planFormatCells(
      req({ target: { range: 'A1' }, format: { bold: false } }, 'format-cells'),
    );
    expect(plan.bold).toBe(false);
    expect(plan.hasOps).toBe(true);
  });
});

describe('planAddComment (ADR-0004 add-comment, Excel)', () => {
  it('shapes the comment to the anchor range and single-lines the text', () => {
    const plan = planAddComment(
      req({ target: { range: 'Sheet1!B2' }, text: 'Spike   here\nplease' }, 'add-comment'),
    );
    expect(plan).toEqual({
      address: 'Sheet1!B2',
      text: 'Spike here please',
      hasTarget: true,
      hasText: true,
    });
  });

  it('reports hasTarget:false when no range is given', () => {
    const plan = planAddComment(req({ text: 'note' }, 'add-comment'));
    expect(plan.hasTarget).toBe(false);
    expect(plan.address).toBeUndefined();
  });

  it('reports hasText:false for empty/whitespace text', () => {
    expect(planAddComment(req({ target: { range: 'A1' } }, 'add-comment')).hasText).toBe(false);
    expect(
      planAddComment(req({ target: { range: 'A1' }, text: '   ' }, 'add-comment')).hasText,
    ).toBe(false);
  });
});
