/**
 * @ge/runtime — the surface-agnostic core. Build the assist loop once; every bridge
 * (Word/Excel/Outlook/…) plugs into it. See docs/CAPABILITY-MAP.md and ADR-0002.
 */
export type { DocBridge, AuthClient, UserIdentity } from './bridge.js';
export { AssistSession, type AssistSessionOptions } from './assist-session.js';
export {
  ContextModel,
  BRIEF_REF_ID,
  type CommitMode,
  type CommitHint,
  type ContextBrief,
} from './context-model.js';
export {
  Orchestrator,
  type OrchestratorHandlers,
  type OrchestratorOptions,
} from './orchestrator.js';
