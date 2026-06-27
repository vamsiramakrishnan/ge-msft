/**
 * @ge/bridge-onenote — the OneNote DocBridge (web-only, legacy XML manifest). Native context
 * capture (active page title + outline rich text) + citation-tagged page synthesis
 * (`append-page`). Implements @ge/runtime's DocBridge, so it plugs into the shared AssistSession
 * loop unchanged.
 */
export { OneNoteBridge, HANDLED_ACTUATIONS } from './onenote-bridge.js';
export { ONENOTE_CAPABILITIES } from './capabilities.js';
export {
  pageElementToBlocks,
  pageElementToDocStateBlocks,
  pageToContext,
  searchPage,
  MAX_SEARCH_PARAGRAPHS,
  type PageElement,
} from './capture.js';
export {
  escapeHtml,
  citationTag,
  partToHtml,
  buildPageHtml,
  type SynthesisPart,
} from './synthesis.js';
export { planAppendPage, type AppendPagePlan } from './actuate-plan.js';
