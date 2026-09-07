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
  type PlanEffect,
  type RunCommandsOptions,
  DOC_STATE_REF_ID,
  READ_REF_PREFIX,
} from './assist-session.js';
export {
  compileCommand,
  isCompileError,
  renderGrammarPrompt,
  renderCommandBootstrap,
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
export { analyseEffectDependencies, propagateFailure, effectResources } from './planning.js';
// Type-only; no runtime cost. @ge/runtime has no package.json subpath exports, so this barrel
// is the only way web-shell/controller.ts's artifact rendering can reach these workspace types.
export { type WorkspaceArtifactSummary, type WorkspaceResult } from './workspace.js';
export * from './docfs/index.js';
export {
  CommandCapsule,
  CommandCapsuleBudgetError,
  type CommandCapsuleOptions,
  type CommandCapsuleTurn,
} from './command-capsule.js';
export {
  RuntimeHooks,
  HookBlockedError,
  type RuntimeHook,
  type RuntimeHookPhase,
  type RuntimeHookPayloads,
  type HookContext,
  type HookRecord,
  type HookDecision,
  type HookResult,
} from './hooks.js';
export {
  ExecutionLedger,
  type RunOutcome,
  type RunRecord,
  type RunStatus,
  type TaskMode,
} from './execution-ledger.js';
export {
  installRuntimeExtensions,
  completedEffectsExtension,
  type RuntimeExtension,
  type RuntimeExtensionApi,
} from './extensions.js';

export * from './analysis-workspace.js';
export * from './recovery.js';
export * from './evidence.js';
export * from './analysis-program.js';
export * from './result-store.js';
export { discoverCommands, type CommandCard } from './capability-catalog.js';
