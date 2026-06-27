/**
 * @ge/gemini-client — client-direct access to Gemini Enterprise (Discovery Engine).
 * The add-in federates the user's Entra identity to Google (WIF) and calls
 * `:streamAssist` directly; the engine owns grounding, Model Armor, and agents.
 * See docs/ADR-0001-client-direct-architecture.md.
 */
export * from './config.js';
export * from './wif.js';
export * from './stream-assist.js';
export * from './session-context.js';
export * from './context-files.js';
export * from './resolve-grounding.js';
export * from './search.js';
export * from './autocomplete.js';
export * from './grounding.js';
export * from './ranking.js';
export * from './catalog.js';
export {
  defaultFetch,
  getJson,
  postJson,
  postJsonWithHeaders,
  type FetchLike,
} from './de-fetch.js';
export { parseJsonArrayStream } from './json-stream.js';
export { contentHash } from './hash.js';
export {
  withRetry,
  backoffDelay,
  defaultIsRetriable,
  HttpError,
  CircuitBreaker,
  type RetryOptions,
  type CircuitState,
} from './retry.js';
export type { DeStreamAssistResponse } from './de-types.js';
export type {
  DeSearchResponse,
  DeCompleteQueryResponse,
  DeCheckGroundingResponse,
  DeRankResponse,
} from './de-search-types.js';
