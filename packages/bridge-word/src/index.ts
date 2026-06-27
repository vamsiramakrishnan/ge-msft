/**
 * @ge/bridge-word — the Word DocBridge. Native context capture (selection / paragraphs /
 * styles) + content-anchored tracked-change writes. Implements @ge/runtime's DocBridge,
 * so it plugs into the shared AssistSession loop unchanged.
 */
export { WordBridge, HANDLED_ACTUATIONS } from './word-bridge.js';
export { WORD_CAPABILITIES } from './capabilities.js';
export {
  wordElementsToBlocks,
  wordDocumentToContext,
  wordSelectionToContext,
  headingLevel,
  type WordElement,
} from './capture.js';
export {
  planTrackedChange,
  planAddComment,
  chooseAnchorIndex,
  formatSources,
  type TrackedChangePlan,
  type AddCommentPlan,
} from './actuate-plan.js';
export { provenanceRecord, provenanceKey, type ProvenanceRecord } from './provenance-record.js';
export {
  originFromWordSource,
  selectionChangedEvent,
  documentChangedEvent,
  commentAddedEvent,
} from './events.js';
export { OfficeWordHost } from './host-port.js';
export type {
  WordHost,
  WordHandlers,
  WordParagraph,
  WordEditArgs,
  WordCommentArgs,
  ChooseHit,
  TrackedChangeOutcome,
  CommentReplyOutcome,
} from './host-port.js';
