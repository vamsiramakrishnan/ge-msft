/**
 * @ge/bridge-powerpoint — the PowerPoint DocBridge. Native context capture (selected slides →
 * shapes' text + speaker notes via the `native.slide()` builder) + deck-composer writes
 * (`insert-slide`, `set-speaker-notes`). Implements @ge/runtime's DocBridge, so it plugs into
 * the shared AssistSession loop unchanged.
 */
export { PowerPointBridge, HANDLED_ACTUATIONS, MAX_READ_SLIDES } from './powerpoint-bridge.js';
export { POWERPOINT_CAPABILITIES } from './capabilities.js';
export {
  shapesToSlideText,
  slideElementsToBlocks,
  slideElementsToDocStateBlocks,
  slidesToContext,
  selectedSlideToContext,
  searchSlides,
  parseSlideSelector,
  MAX_SEARCH_SLIDES,
  type SlideElement,
} from './capture.js';
export {
  planInsertSlide,
  planSpeakerNotes,
  type InsertSlidePlan,
  type SpeakerNotesPlan,
} from './actuate-plan.js';
export { selectionChanged, documentChanged } from './events.js';
