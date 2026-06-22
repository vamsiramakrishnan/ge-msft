import { describe, it, expect } from 'vitest';
import { ResolvedContextSchema, asChangeId } from '@ge/contracts';
import {
  headingLevel,
  wordDocumentToContext,
  wordElementsToBlocks,
  wordSelectionToContext,
  type WordElement,
} from './capture.js';
import { chooseAnchorIndex, planTrackedChange } from './actuate-plan.js';

describe('word capture (pure)', () => {
  it('derives heading levels from built-in styles', () => {
    expect(headingLevel('Heading1')).toBe(1);
    expect(headingLevel('Heading 2')).toBe(2);
    expect(headingLevel('Normal')).toBe(0);
  });

  it('maps native elements to blocks with cc: locators', () => {
    const els: WordElement[] = [
      { kind: 'heading', text: '5. Service Levels', level: 1, contentControlId: 10 },
      { kind: 'paragraph', text: 'Available 99.5% of the time.', contentControlId: 12 },
    ];
    const blocks = wordElementsToBlocks(els);
    expect(blocks[0]).toMatchObject({ kind: 'heading', locator: 'cc:10' });
    expect(blocks[1]).toMatchObject({ kind: 'paragraph', locator: 'cc:12' });
  });

  it('produces valid, anchored context from a document', () => {
    const ctx = wordDocumentToContext(
      'word:document',
      'MSA',
      [
        { kind: 'heading', text: 'Availability', level: 2, contentControlId: 7 },
        {
          kind: 'paragraph',
          text: 'The services are available 99.5% of the time.',
          contentControlId: 8,
        },
      ],
      { maxTokens: 40 },
    );
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator?.startsWith('cc:'))).toBe(true);
  });

  it('turns a selection into a single live text part (and skips empty)', () => {
    expect(wordSelectionToContext('  ')).toHaveLength(0);
    const ctx = wordSelectionToContext('hello');
    expect(ctx[0]).toMatchObject({ ref: { kind: 'selection', live: true }, value: { as: 'text' } });
  });
});

describe('word actuation planning (pure)', () => {
  it('plans a content-anchored tracked change', () => {
    const plan = planTrackedChange({
      changeId: asChangeId('c1'),
      kind: 'tracked-change',
      surface: 'word',
      params: {
        text: '99.9%',
        target: { matchText: '99.5%', contextHint: 'Availability' },
      },
    });
    expect(plan).toEqual({ matchText: '99.5%', contextHint: 'Availability', text: '99.9%' });
  });

  it('chooses the anchor hit matching the contextHint, else the first', () => {
    expect(chooseAnchorIndex(['intro 99.5%', 'Availability: 99.5%'], 'Availability')).toBe(1);
    expect(chooseAnchorIndex(['a', 'b'])).toBe(0);
    expect(chooseAnchorIndex([])).toBe(-1);
  });
});
