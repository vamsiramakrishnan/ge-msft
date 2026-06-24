import { describe, it, expect } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import {
  planFillContentControl,
  planInsertOoxml,
  planInsertText,
  planReplaceSelection,
} from './actuate-plan.js';

function req(kind: ActuationRequest['kind'], params: ActuationRequest['params']): ActuationRequest {
  return { changeId: asChangeId('c1'), kind, surface: 'word', params };
}

describe('planInsertText (ADR-0007 insert-text)', () => {
  it('marks anchored=true and carries matchText/contextHint when an anchor is given', () => {
    const plan = planInsertText(
      req('insert-text', {
        text: 'hi',
        target: { matchText: '99.5%', contextHint: 'Availability' },
      }),
    );
    expect(plan).toEqual({
      matchText: '99.5%',
      contextHint: 'Availability',
      anchored: true,
      text: 'hi',
      hasText: true,
    });
  });

  it('marks anchored=false and omits matchText when no anchor is given (selection path)', () => {
    const plan = planInsertText(req('insert-text', { text: 'hi' }));
    expect(plan.anchored).toBe(false);
    expect(plan.matchText).toBeUndefined();
    expect(plan.hasText).toBe(true);
  });

  it('reports hasText=false for empty/absent text', () => {
    expect(planInsertText(req('insert-text', {})).hasText).toBe(false);
    expect(planInsertText(req('insert-text', { text: '' })).hasText).toBe(false);
  });
});

describe('planReplaceSelection (ADR-0007 replace-selection)', () => {
  it('carries text and hasText', () => {
    expect(planReplaceSelection(req('replace-selection', { text: 'new' }))).toEqual({
      text: 'new',
      hasText: true,
    });
  });
  it('reports hasText=false for empty/absent text', () => {
    expect(planReplaceSelection(req('replace-selection', {})).hasText).toBe(false);
    expect(planReplaceSelection(req('replace-selection', { text: '' })).hasText).toBe(false);
  });
});

describe('planInsertOoxml (ADR-0007 insert-ooxml)', () => {
  it('marks anchored=true with the anchor and carries the ooxml', () => {
    const plan = planInsertOoxml(
      req('insert-ooxml', { ooxml: '<w:p/>', target: { matchText: 'Summary' } }),
    );
    expect(plan).toEqual({
      matchText: 'Summary',
      anchored: true,
      ooxml: '<w:p/>',
      hasOoxml: true,
    });
  });
  it('marks anchored=false at the selection and reports hasOoxml=false for empty ooxml', () => {
    expect(planInsertOoxml(req('insert-ooxml', { ooxml: '<w:p/>' })).anchored).toBe(false);
    expect(planInsertOoxml(req('insert-ooxml', { ooxml: '' })).hasOoxml).toBe(false);
    expect(planInsertOoxml(req('insert-ooxml', {})).hasOoxml).toBe(false);
  });
});

describe('planFillContentControl (ADR-0007 fill-content-control)', () => {
  it('carries the id + text with hasId/hasText', () => {
    const plan = planFillContentControl(
      req('fill-content-control', { target: { contentControlId: '42' }, text: 'filled' }),
    );
    expect(plan).toEqual({
      contentControlId: '42',
      hasId: true,
      text: 'filled',
      hasText: true,
    });
  });
  it('reports hasId=false when no contentControlId is given', () => {
    const plan = planFillContentControl(req('fill-content-control', { text: 'x' }));
    expect(plan.hasId).toBe(false);
    expect(plan.contentControlId).toBeUndefined();
  });
  it('reports hasText=false for empty text', () => {
    const plan = planFillContentControl(
      req('fill-content-control', { target: { contentControlId: '42' }, text: '' }),
    );
    expect(plan.hasText).toBe(false);
  });
});
