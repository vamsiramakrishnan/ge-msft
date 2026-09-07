import type { ActuationResult, Surface } from '@ge/contracts';

export type TaskMode = 'chat' | 'command' | 'program' | 'planner' | 'proposal';
export type RunStatus = 'running' | 'completed' | 'failed' | 'blocked' | 'cancelled' | 'incomplete';
/** Full outcomes are ephemeral hook inputs, including inverse receipts when the host supplies them. */
export interface RunOutcome {
  taskId: string;
  surface: Surface;
  mode: TaskMode;
  status: RunStatus;
  startedAt: string;
  modelTurns: number;
  toolCalls: number;
  effects: ActuationResult[];
}
/** The diagnostic ledger deliberately excludes prompts, source content, model output, and inverse bodies. */
export interface RunRecord extends Omit<RunOutcome, 'effects'> {
  finishedAt?: string;
  effects: Array<{
    changeId: string;
    kind: ActuationResult['kind'];
    ok: boolean;
    errorCode?: string;
  }>;
}
export class ExecutionLedger {
  private readonly runs: RunRecord[] = [];
  private readonly listeners = new Set<(record: RunRecord) => void>();
  record(outcome: RunOutcome): void {
    const record: RunRecord = {
      taskId: outcome.taskId,
      surface: outcome.surface,
      mode: outcome.mode,
      status: outcome.status,
      startedAt: outcome.startedAt,
      modelTurns: outcome.modelTurns,
      toolCalls: outcome.toolCalls,
      ...(outcome.status !== 'running' ? { finishedAt: new Date().toISOString() } : {}),
      effects: outcome.effects.map((r) => ({
        changeId: r.changeId,
        kind: r.kind,
        ok: r.ok,
        ...(r.error ? { errorCode: r.error.code } : {}),
      })),
    };
    const index = this.runs.findIndex((r) => r.taskId === outcome.taskId);
    if (index < 0) this.runs.push(record);
    else this.runs[index] = record;
    if (this.runs.length > 100) this.runs.shift();
    for (const listener of this.listeners) {
      try {
        listener(structuredClone(record));
      } catch {
        /* observational */
      }
    }
  }
  list(): RunRecord[] {
    return structuredClone(this.runs);
  }
  subscribe(listener: (record: RunRecord) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}
