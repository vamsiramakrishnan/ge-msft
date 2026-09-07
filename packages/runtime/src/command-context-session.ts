import type { ActuationResult, ArtifactSummary, CapabilityManifest } from '@ge/contracts';
import { renderCommandBootstrap, renderGrammarPrompt } from './command-protocol.js';
import {
  CommandCapsule,
  type CommandCapsuleSkill,
  type CommandCapsuleTurn,
} from './command-capsule.js';
import { ExecutionState, type ExecutionStateBinding } from './execution-state.js';
import { CommandResultStore } from './result-store.js';
import type { Value } from './compose.js';

export interface CommandContextSessionOptions {
  sessionMode?: 'sessionless' | 'conversation';
  contextMode?: 'projection' | 'transcript';
  disclosure?: 'compact' | 'full';
  maxBytes?: number;
}
export type CommandObservation =
  | { program: string; results: readonly unknown[]; correction?: never }
  | { program: string; correction: string; results?: never };
/** Existing task counters; recording updates them once, even if subsequent journal retention fails. */
export interface CommandEncodingMetrics {
  resultInputBytes: number;
  resultInputBytesComplete?: boolean;
  resultOutputBytes: number;
}
export interface CommandContextRuntimeState {
  analysisBindings: readonly (readonly [string, string])[];
  composeBindings: ReadonlyMap<string, Value>;
  artifacts: readonly ArtifactSummary[];
  effects: readonly ActuationResult[];
  externalShareAttempts: number;
}
export interface CommandContextRenderInput {
  capabilities: CapabilityManifest;
  docState?: string;
  skills?: readonly CommandCapsuleSkill[];
  /** Read only when projection needs live state; transcript/conversation avoid collecting artifacts. */
  state?: () => CommandContextRuntimeState;
}

/**
 * One owner for command observation retention, disclosure policy and inspection lifetime. The
 * representation never executes commands, captures host state, grants approval, or updates the
 * provider's conversation. Those responsibilities remain with AssistSession.
 */
export class CommandContextSession {
  private readonly results = new CommandResultStore({ inlineBytes: 4096 });
  private history?: CommandCapsule | ExecutionState;
  private latest?: CommandCapsuleTurn;
  private task?: string;

  private readonly options: Readonly<CommandContextSessionOptions>;
  constructor(options: CommandContextSessionOptions = {}) {
    this.options = { ...options };
  }

  get isolated(): boolean {
    return this.options.sessionMode !== 'conversation';
  }

  /** The caller captures host state using its existing freshness/hook policy. */
  get snapshotPolicy(): 'fresh' | 'deduplicate' | 'none' {
    if (this.isolated) return 'fresh';
    return this.latest?.correction !== undefined ? 'none' : 'deduplicate';
  }

  begin(task: string): void {
    this.clear();
    if (this.isolated)
      this.history =
        this.options.contextMode === 'transcript'
          ? new CommandCapsule(task, { maxBytes: this.options.maxBytes })
          : new ExecutionState(task, { maxBytes: this.options.maxBytes });
    this.task = task;
  }

  clear(): void {
    this.results.clear();
    if (this.history instanceof ExecutionState) this.history.clear();
    this.history = undefined;
    this.latest = undefined;
    this.task = undefined;
  }

  record(observation: CommandObservation, metrics?: CommandEncodingMetrics): void {
    this.requireTask();
    let resultsJson: string | undefined;
    if (observation.results !== undefined) {
      const encoded = this.results.encode(observation.results);
      resultsJson = encoded.text;
      if (metrics) {
        metrics.resultInputBytes += encoded.inputBytes;
        metrics.resultInputBytesComplete =
          metrics.resultInputBytesComplete !== false && encoded.inputBytesComplete;
        metrics.resultOutputBytes += encoded.outputBytes;
      }
    }
    const turn: CommandCapsuleTurn = {
      program: observation.program,
      ...(resultsJson !== undefined ? { resultsJson } : {}),
      ...(observation.correction !== undefined ? { correction: observation.correction } : {}),
    };
    if (this.history instanceof ExecutionState) this.history.append(turn, observation.results);
    else this.history?.append(turn);
    this.latest = turn;
  }

  render(input: CommandContextRenderInput): string {
    const task = this.requireTask();
    if (!this.history && this.latest) {
      if (this.latest.correction !== undefined) return this.latest.correction;
      const resultBlock = `\`\`\`result\n${this.latest.resultsJson}\n\`\`\``;
      return input.docState
        ? `${resultBlock}\n\n${input.docState}\n\n(Continue. Next command?)`
        : `${resultBlock}\n\n(Continue. Next command?)`;
    }
    const protocol =
      this.options.disclosure === 'full'
        ? renderGrammarPrompt(input.capabilities)
        : renderCommandBootstrap(input.capabilities, task);
    if (this.history) {
      const state = this.history instanceof ExecutionState ? input.state?.() : undefined;
      return this.history.render({
        ...(state ? projectState(state) : {}),
        protocol,
        docState: input.docState,
        skills: input.skills,
      });
    }
    const parts = [protocol];
    if (input.docState) parts.push(input.docState);
    parts.push(`TASK:\n${task}`, 'Begin.');
    return parts.join('\n\n');
  }

  /** Undefined means this selector belongs to a host/workspace namespace, so its reader may run. */
  inspect(selector: string): unknown | undefined {
    const selected = selector.trim();
    if (selected.startsWith('result:')) return this.results.inspect(selected);
    if (selected.startsWith('state:'))
      return this.history instanceof ExecutionState
        ? this.history.inspect(selected)
        : { error: 'Execution state reference expired.', code: 'reference', complete: false };
    return undefined;
  }

  private requireTask(): string {
    if (this.task === undefined)
      throw new Error('Begin a command context before recording or rendering.');
    return this.task;
  }
}

function projectState(state: CommandContextRuntimeState) {
  const artifacts = state.artifacts.map(
    ({ preview: _preview, createdAt: _createdAt, ...artifact }) => artifact,
  );
  const available = new Set(artifacts.map((artifact) => artifact.id));
  const bindings: ExecutionStateBinding[] = state.analysisBindings.map(([name, id]) => ({
    name,
    kind: 'artifact',
    value: { id, available: available.has(id) },
  }));
  for (const [name, value] of state.composeBindings)
    bindings.push({
      name,
      kind: value.kind,
      value,
      ...(value.kind === 'table'
        ? { schema: { columns: value.columns, rows: value.rows.length } }
        : {}),
    });
  return {
    bindings,
    artifacts,
    effects: state.effects,
    externalShareAttempts: state.externalShareAttempts,
  };
}
