import { describe, it, expect } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { planInsertSlide, planSpeakerNotes } from './actuate-plan.js';

describe('powerpoint insert-slide planning (pure)', () => {
  it('extracts title, bullets, notes, and target index', () => {
    const plan = planInsertSlide({
      changeId: asChangeId('c1'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: {
        slide: { title: 'SLA', bullets: ['99.5% contracted', 'gap flagged'], notes: 'Open here' },
        target: { slideIndex: 2 },
      },
    });
    expect(plan).toEqual({
      title: 'SLA',
      bullets: ['99.5% contracted', 'gap flagged'],
      notes: 'Open here',
      targetIndex: 2,
    });
  });

  it('prefers a prebuilt base64 deck when provided via ooxml', () => {
    const plan = planInsertSlide({
      changeId: asChangeId('c2'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: { ooxml: 'UEsDBBQ=' },
    });
    expect(plan.base64).toBe('UEsDBBQ=');
    expect(plan.title).toBe('');
    expect(plan.bullets).toEqual([]);
  });

  it('extracts an explicit prebuilt deck artifact before legacy ooxml', () => {
    const plan = planInsertSlide({
      changeId: asChangeId('c-deck'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: {
        ooxml: 'legacy',
        deck: {
          format: 'pptx',
          base64: 'compiled',
          slideCount: 4,
          formatting: 'UseDestinationTheme',
          targetSlideId: '256#3',
          specFingerprint: 'abc123',
        },
      },
    });
    expect(plan).toMatchObject({
      base64: 'compiled',
      slideCount: 4,
      formatting: 'UseDestinationTheme',
      targetSlideId: '256#3',
      specFingerprint: 'abc123',
    });
  });

  it('omits notes when blank and target when absent', () => {
    const plan = planInsertSlide({
      changeId: asChangeId('c3'),
      kind: 'insert-slide',
      surface: 'powerpoint',
      params: { slide: { title: 'T', bullets: [], notes: '   ' } },
    });
    expect(plan).toEqual({ title: 'T', bullets: [] });
  });
});

describe('powerpoint set-speaker-notes planning (pure)', () => {
  it('takes notes from params.slide.notes and a targeted slide index', () => {
    const plan = planSpeakerNotes({
      changeId: asChangeId('c4'),
      kind: 'set-speaker-notes',
      surface: 'powerpoint',
      params: { slide: { title: '', bullets: [], notes: 'Talk track' }, target: { slideIndex: 1 } },
    });
    expect(plan).toEqual({ notes: 'Talk track', targetIndex: 1 });
  });

  it('falls back to params.text and omits the index when untargeted', () => {
    const plan = planSpeakerNotes({
      changeId: asChangeId('c5'),
      kind: 'set-speaker-notes',
      surface: 'powerpoint',
      params: { text: 'From text' },
    });
    expect(plan).toEqual({ notes: 'From text' });
  });
});
