import { describe, expect, it } from 'vitest';
import { asChangeId } from '@ge/contracts';
import type { PlanEffect } from '../../controller.js';
import { effectChange, isHighImpact } from './PlanApprovalCard.js';

const reply = (attachment?: PlanEffect['request']['params']['attachment']): PlanEffect => ({
  request: {
    changeId: asChangeId('c-1'),
    kind: 'reply-mail',
    surface: 'outlook',
    params: { mail: { body: 'See attached.' }, ...(attachment ? { attachment } : {}) },
  },
  command: 'mail',
  approvalClass: 'external',
  reversible: true,
});

describe('approval card: a reply that carries a file (security review 2026-10-17)', () => {
  it('names the file before the body and flags the row for review', () => {
    const effect = reply({ name: 'notes.docx', base64: 'UEsDBA==', text: 'secret figures' });
    expect(effectChange(effect)).toBe('file: notes.docx · body: See attached.');
    expect(isHighImpact(effect)).toBe(true);
  });

  it('leaves a plain reply unflagged', () => {
    expect(effectChange(reply())).toBe('body: See attached.');
    expect(isHighImpact(reply())).toBe(false);
  });
});
