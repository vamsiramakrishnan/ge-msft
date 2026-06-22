/**
 * @ge/runtime — the surface-agnostic core. Build the assist loop once; every bridge
 * (Word/Excel/Outlook/…) plugs into it. See docs/CAPABILITY-MAP.md and ADR-0002.
 */
export type { DocBridge, AuthClient, UserIdentity } from './bridge.js';
export {
  AssistSession,
  type AssistSessionOptions,
  type CompactionOptions,
  type ContextLoopOptions,
  type CommandLoopEvent,
  type RunCommandsOptions,
  DOC_STATE_REF_ID,
  READ_REF_PREFIX,
} from './assist-session.js';
export {
  compileCommand,
  isCompileError,
  renderGrammarPrompt,
  type CompiledCommand,
  type ReadIntent,
} from './command-protocol.js';
export {
  TRANSFORMS,
  TRANSFORM_NAMES,
  TRANSFORM_USAGE,
  parseTable,
  renderValue,
  evalExpr,
  isEvalError,
  type Value,
  type EvalError,
  type Transform,
  type RunRead,
} from './compose.js';
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
