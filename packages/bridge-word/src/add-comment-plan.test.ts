import { describe, it, expect } from 'vitest';
import { asChangeId, type ActuationRequest } from '@ge/contracts';
import { planAddComment } from './actuate-plan.js';

function addComment(params: ActuationRequest['params']): ActuationRequest {
  return { changeId: asChangeId('c1'), kind: 'add-comment', surface: 'word', params };
}

describe('planAddComment (ADR-0004 add-comment, Word)', () => {
  it('anchors by matchText + contextHint and single-lines the comment text', () => {
    const plan = planAddComment(
      addComment({
        target: { matchText: '99.5%', contextHint: 'Availability' },
        text: 'Unsourced   claim\nflag this',
      }),
    );
    expect(plan).toEqual({
      matchText: '99.5%',
      contextHint: 'Availability',
      text: 'Unsourced claim flag this',
      hasText: true,
    });
  });

  it('omits matchText/contextHint when absent', () => {
    const plan = planAddComment(addComment({ text: 'note' }));
    expect(plan.matchText).toBeUndefined();
    expect(plan.contextHint).toBeUndefined();
    expect(plan.hasText).toBe(true);
  });

  it('reports hasText:false for empty/whitespace text', () => {
    expect(planAddComment(addComment({ target: { matchText: 'x' } })).hasText).toBe(false);
    expect(planAddComment(addComment({ target: { matchText: 'x' }, text: '  ' })).hasText).toBe(
      false,
    );
  });
});
