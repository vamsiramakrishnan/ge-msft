/**
 * @ge/bridge-word — the Word DocBridge. Native context capture (selection / paragraphs /
 * styles) + content-anchored tracked-change writes. Implements @ge/runtime's DocBridge,
 * so it plugs into the shared AssistSession loop unchanged.
 */
export { WordBridge } from './word-bridge.js';
export { WORD_CAPABILITIES } from './capabilities.js';
export {
  wordElementsToBlocks,
  wordDocumentToContext,
  wordSelectionToContext,
  headingLevel,
  type WordElement,
} from './capture.js';
export { planTrackedChange, chooseAnchorIndex, type TrackedChangePlan } from './actuate-plan.js';
export {
  originFromWordSource,
  selectionChangedEvent,
  documentChangedEvent,
  commentAddedEvent,
} from './events.js';
