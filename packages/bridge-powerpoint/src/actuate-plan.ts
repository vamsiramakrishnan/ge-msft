import type { ActuationRequest } from '@ge/contracts';

/**
 * Pure translation of a PowerPoint actuation into a host plan — testable without Office.js.
 *
 * `insert-slide`: a slide is composed from `params.slide` ({ title, bullets, notes }). When the
 * agent supplies a prebuilt PPTX deck as base64 (`params.deck.base64`, with legacy `params.ooxml`
 * still accepted), the bridge prefers `insertSlidesFromBase64`; otherwise it appends a native slide
 * and writes the placeholders.
 *
 * `set-speaker-notes`: notes text targeted at a slide (`params.target.slideIndex`, else the
 * selected/active slide). The notes string comes from `params.slide.notes` or `params.text`.
 */

export interface InsertSlidePlan {
  title: string;
  bullets: string[];
  notes?: string;
  /** Prebuilt single/multi-slide deck as a Base64 PPTX, when the agent provides one. */
  base64?: string;
  slideCount?: number;
  formatting?: 'KeepSourceFormatting' | 'UseDestinationTheme';
  targetSlideId?: string;
  specFingerprint?: string;
  /** Zero-based insertion index, when targeted; otherwise append to the end. */
  targetIndex?: number;
}

export function planInsertSlide(req: ActuationRequest): InsertSlidePlan {
  const p = req.params;
  const slide = p.slide;
  const plan: InsertSlidePlan = {
    title: slide?.title ?? '',
    bullets: slide?.bullets ?? [],
  };
  const notes = slide?.notes ?? p.text;
  if (notes !== undefined && notes.trim().length > 0) plan.notes = notes;
  const deck = p.deck;
  const base64 = deck?.base64 ?? p.ooxml;
  if (base64 !== undefined && base64.length > 0) plan.base64 = base64;
  if (deck?.slideCount !== undefined) plan.slideCount = deck.slideCount;
  if (deck?.formatting !== undefined) plan.formatting = deck.formatting;
  if (deck?.targetSlideId !== undefined) plan.targetSlideId = deck.targetSlideId;
  if (deck?.specFingerprint !== undefined) plan.specFingerprint = deck.specFingerprint;
  if (p.target?.slideId !== undefined && plan.targetSlideId === undefined) {
    plan.targetSlideId = p.target.slideId;
  }
  if (typeof p.target?.slideIndex === 'number') plan.targetIndex = p.target.slideIndex;
  return plan;
}

export interface SpeakerNotesPlan {
  notes: string;
  /** Zero-based slide index to target; undefined ⇒ the selected/active slide. */
  targetIndex?: number;
}

export function planSpeakerNotes(req: ActuationRequest): SpeakerNotesPlan {
  const p = req.params;
  const plan: SpeakerNotesPlan = { notes: p.slide?.notes ?? p.text ?? '' };
  if (typeof p.target?.slideIndex === 'number') plan.targetIndex = p.target.slideIndex;
  return plan;
}
