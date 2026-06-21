/**
 * @ge/teams — the Teams DocBridge. String-path context capture (a meeting/chat transcript window,
 * labelled and normalized via @ge/content) + reviewable, provenanced chat posts (`post-message`)
 * staged through a feature-detected TeamsJS surface. Implements @ge/runtime's DocBridge, so it
 * plugs into the shared AssistSession loop unchanged.
 */
export {
  TeamsBridge,
  type TeamsBridgeOptions,
  type TeamsJsLike,
  type TeamsComposeRequest,
} from './teams-bridge.js';
export { TEAMS_CAPABILITIES } from './capabilities.js';
export { transcriptToContext, type TranscriptInput } from './capture.js';
export { planPostMessage, type PostMessagePlan } from './actuate-plan.js';
export { sessionStartEvent, sessionEndEvent, meetingEndedEvent } from './events.js';
