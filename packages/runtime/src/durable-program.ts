import { z } from 'zod';
import {
  ActuationResultSchema,
  ActuationRequestSchema,
  ActuationKindSchema,
  DurableCommandProgramSchema,
  assessActuationResult,
  validateActuationResult,
  asChangeId,
  contentHash,
  parseProgramBlock,
  isProgramSkillCall,
  isProgramSkillDef,
  type ActuationResult,
  type ActuationRequest,
  type DurableCommandProgram,
  type DurableProgramStage,
} from '@ge/contracts';

const HASH = /^sha256:[a-f0-9]{64}$/;
const CheckpointSchema = z
  .object({
    id: z.string().min(1).max(64),
    state: z.enum(['pending', 'running', 'verified', 'blocked']),
    startedAt: z.string().datetime().optional(),
    effect: z
      .object({
        changeId: z.string().min(1).max(256),
        kind: ActuationKindSchema,
        fingerprint: z.string().regex(HASH),
      })
      .strict()
      .optional(),
    inspections: z
      .array(
        z
          .object({
            at: z.string().datetime(),
            decision: z.enum(['not-applied', 'verified', 'unresolved']),
            priorState: z.enum(['running', 'blocked']),
            changeId: z.string().max(256).optional(),
          })
          .strict(),
      )
      .max(16)
      .default([]),
    receipt: ActuationResultSchema.optional(),
  })
  .strict();
const RecordSchema = z
  .object({
    version: z.literal(1),
    program: DurableCommandProgramSchema,
    fingerprint: z.string().regex(HASH),
    updatedAt: z.string().datetime(),
    runId: z.string().uuid().optional(),
    stages: z.array(CheckpointSchema).max(32),
  })
  .strict();
export type DurableProgramRecord = z.infer<typeof RecordSchema>;
export type DurableStageCheckpoint = z.infer<typeof CheckpointSchema>;
const HistorySchema = z.array(RecordSchema).max(32);
function history(raw: unknown): DurableProgramRecord[] {
  const records = HistorySchema.parse(raw === undefined ? [] : raw);
  if (new Set(records.map((record) => record.program.id)).size !== records.length)
    throw new Error('Duplicate durable program IDs.');
  for (const record of records) {
    if (
      record.stages.length !== record.program.stages.length ||
      record.stages.some(
        (stage, index) =>
          stage.id !== record.program.stages[index]?.id ||
          (stage.state === 'verified' &&
            (!stage.effect ||
              !stage.receipt ||
              stage.receipt.changeId !== stage.effect.changeId ||
              stage.receipt.kind !== stage.effect.kind ||
              assessActuationResult(stage.receipt) !== 'verified')) ||
          (stage.state === 'pending' &&
            stage.receipt &&
            !['failed', 'rejected'].includes(assessActuationResult(stage.receipt))),
      )
    )
      throw new Error('Durable checkpoints do not match their program or verified receipts.');
  }
  if (new TextEncoder().encode(JSON.stringify(records)).byteLength > 8 * 1024 * 1024)
    throw new Error('Durable program storage exceeds 8 MiB. Archive completed tasks explicitly.');
  return records;
}
/** update must be atomic across panes. No get/set emulation is safe for claims. */
export interface DurableProgramStore {
  readonly durable: boolean;
  read(scope: string): Promise<unknown>;
  update(scope: string, change: (records: unknown) => DurableProgramRecord[]): Promise<void>;
}
export class MemoryDurableProgramStore implements DurableProgramStore {
  readonly durable = false;
  private readonly values = new Map<string, DurableProgramRecord[]>();
  async read(scope: string): Promise<unknown> {
    return structuredClone(this.values.get(scope) ?? []);
  }
  async update(scope: string, change: (records: unknown) => DurableProgramRecord[]): Promise<void> {
    this.values.set(
      scope,
      structuredClone(history(change(structuredClone(this.values.get(scope) ?? [])))),
    );
  }
}
/** Native IndexedDB transactions serialize claims across task panes and survive reload. */
export class IndexedDbDurableProgramStore implements DurableProgramStore {
  readonly durable = true;
  private readonly database: Promise<IDBDatabase>;
  constructor(factory: IDBFactory = indexedDB) {
    this.database = new Promise((resolve, reject) => {
      const request = factory.open('ge-msft-durable-programs', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('programs');
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onerror = () => reject(new Error('Durable program storage could not open.'));
      request.onblocked = () => reject(new Error('Close an older pane to upgrade task storage.'));
    });
  }
  async read(scope: string): Promise<unknown> {
    const db = await this.database;
    return new Promise((resolve, reject) => {
      const tx = db.transaction('programs', 'readonly');
      const request = tx.objectStore('programs').get(scope);
      tx.oncomplete = () => resolve(request.result === undefined ? [] : request.result);
      tx.onabort = tx.onerror = () => reject(new Error('Durable program read failed.'));
    });
  }
  async update(scope: string, change: (records: unknown) => DurableProgramRecord[]): Promise<void> {
    const db = await this.database;
    return new Promise((resolve, reject) => {
      const tx = db.transaction('programs', 'readwrite');
      const store = tx.objectStore('programs');
      const request = store.get(scope);
      let failure: unknown;
      request.onsuccess = () => {
        try {
          store.put(history(change(request.result === undefined ? [] : request.result)), scope);
        } catch (error) {
          failure = error;
          tx.abort();
        }
      };
      tx.oncomplete = () => resolve();
      tx.onabort = tx.onerror = () => reject(failure ?? new Error('Durable program write failed.'));
    });
  }
}
export interface DurableProgramEvent {
  type: 'durable-progress';
  programId: string;
  stageId: string;
  state: 'verified' | 'resuming' | 'blocked';
}
export type DurableReconciliationDecision =
  | { decision: 'not-applied' | 'unresolved' }
  | { decision: 'verified'; receipt: ActuationResult };
function checkpointReceipt(raw: ActuationResult): ActuationResult {
  // No inverse, source text, location, free-form error, provenance body or credentials are retained.
  return ActuationResultSchema.parse({
    ok: raw.ok,
    kind: raw.kind,
    changeId: raw.changeId,
    ...(raw.completion ? { completion: raw.completion } : {}),
    ...(raw.verification
      ? {
          verification: {
            status: raw.verification.status,
            ...(HASH.test(raw.verification.afterHash ?? '')
              ? { afterHash: raw.verification.afterHash }
              : {}),
          },
        }
      : {}),
    ...(raw.recoveryPending ? { recoveryPending: true } : {}),
  });
}
function validateProgram(input: DurableCommandProgram): DurableCommandProgram {
  const program = DurableCommandProgramSchema.parse(input);
  for (const stage of program.stages) {
    const parsed = parseProgramBlock('```cmd\n' + stage.program + '\n```');
    if (
      parsed.entries.length === 0 ||
      parsed.entries.some(
        (entry) => 'error' in entry || isProgramSkillCall(entry) || isProgramSkillDef(entry),
      )
    ) {
      throw new Error('Durable stages cannot depend on session macros or malformed commands.');
    }
    if (parsed.entries.some((entry) => 'verb' in entry && entry.verb === 'share'))
      throw new Error('Durable stages cannot contain estate shares.');
  }
  return program;
}
/** Persisted programs are data. This runner never actuates; the session supplies its existing executor. */
export class DurableProgramRunner {
  constructor(
    private readonly store: DurableProgramStore,
    private readonly scope: string,
  ) {
    if (!HASH.test(scope))
      throw new Error('Durable program scope must hash principal and document.');
  }
  async inspect(id: string): Promise<DurableProgramRecord | undefined> {
    const record = history(await this.store.read(this.scope)).find(
      (record) => record.program.id === id,
    );
    if (record && (await contentHash(record.program)) !== record.fingerprint)
      throw new Error('Saved program fingerprint mismatch.');
    return structuredClone(record);
  }
  /** Explicit operator reconciliation after all dispatching panes have stopped. Never a timed retry lease. */
  async reconcile(
    id: string,
    stageId: string,
    inspectCurrent: (
      record: DurableProgramRecord,
      stage: DurableStageCheckpoint,
    ) => Promise<DurableReconciliationDecision>,
    options: { confirmStopped: true },
  ): Promise<void> {
    if (options.confirmStopped !== true)
      throw new Error('Confirm other executing panes have stopped before inspection.');
    const record = await this.inspect(id);
    const stage = record?.stages.find((item) => item.id === stageId);
    if (!record || !stage || !['running', 'blocked'].includes(stage.state))
      throw new Error('Only an interrupted or blocked stage can be reconciled.');
    if (!stage.startedAt || Date.now() - Date.parse(stage.startedAt) < 5 * 60 * 1000)
      throw new Error(
        'This stage may still be in flight. Wait five minutes and stop other panes before inspection.',
      );
    const decision = await inspectCurrent(structuredClone(record), structuredClone(stage));
    if (!['not-applied', 'verified', 'unresolved'].includes(decision.decision))
      throw new Error('Inspection did not establish a supported decision.');
    let receipt: ActuationResult | undefined;
    if (decision.decision === 'verified') {
      if (!stage.effect) throw new Error('There is no correlated effect intent to verify.');
      receipt = validateActuationResult(
        {
          changeId: asChangeId(stage.effect.changeId),
          kind: stage.effect.kind,
          surface: record.program.surface,
          params: {},
        },
        decision.receipt,
      );
      if (
        !stage.effect ||
        receipt.changeId !== stage.effect.changeId ||
        receipt.kind !== stage.effect.kind ||
        assessActuationResult(receipt) !== 'verified'
      )
        throw new Error(
          'Reconciliation requires a correlated native verified receipt and current source/output inspection.',
        );
    } else if (decision.decision === 'not-applied' && stage.receipt?.ok)
      throw new Error(
        'A known applied receipt cannot be reset as not applied. Inspect and verify its output.',
      );
    await this.store.update(this.scope, (raw) =>
      history(raw).map((current) => {
        if (current.program.id !== id) return current;
        const before = current.stages.find((item) => item.id === stageId)!;
        if (
          current.runId !== record.runId ||
          before.state !== stage.state ||
          before.startedAt !== stage.startedAt ||
          before.effect?.fingerprint !== stage.effect?.fingerprint
        )
          throw new Error('The stage changed during inspection. Inspect its latest outcome.');
        const inspection = {
          at: new Date().toISOString(),
          decision: decision.decision,
          priorState: stage.state as 'running' | 'blocked',
          ...(stage.effect ? { changeId: stage.effect.changeId } : {}),
        };
        const inspections = [...before.inspections, inspection];
        if (inspections.length > 16)
          throw new Error('Inspection audit limit reached. Archive this task explicitly.');
        const stages = current.stages.map((item) =>
          item.id !== stageId
            ? item
            : decision.decision === 'not-applied'
              ? { id: item.id, state: 'pending' as const, inspections }
              : {
                  ...item,
                  state:
                    decision.decision === 'verified' ? ('verified' as const) : ('blocked' as const),
                  inspections,
                  ...(receipt ? { receipt: checkpointReceipt(receipt) } : {}),
                },
        );
        const { runId: _runId, ...released } = current;
        return { ...released, stages, updatedAt: new Date().toISOString() };
      }),
    );
  }

  async *run<E>(
    input: DurableCommandProgram,
    execute: (
      stage: DurableProgramStage,
      checkpointEffect: (request: ActuationRequest) => Promise<void>,
      assertOwned: () => Promise<void>,
    ) => AsyncGenerator<E, ActuationResult[]>,
    options: {
      signal?: AbortSignal;
      /** Trusted SDK callback: re-read current source/output state; never persisted as authority. */
      revalidate?: (
        program: DurableCommandProgram,
        checkpoints: DurableStageCheckpoint[],
      ) => Promise<boolean>;
    } = {},
  ): AsyncGenerator<E | DurableProgramEvent> {
    const program = validateProgram(input);
    const fingerprint = await contentHash(program);
    options.signal?.throwIfAborted();
    const runId = crypto.randomUUID();
    await this.store.update(this.scope, (raw) => {
      const records = history(raw);
      const prior = records.find((record) => record.program.id === program.id);
      if (prior && prior.fingerprint !== fingerprint)
        throw new Error('Saved program changed. Use a new task ID and review it.');
      if (prior?.runId)
        throw new Error(
          'Another pane owns this program or its execution was interrupted. Inspect it before proceeding.',
        );
      if (prior) return records.map((record) => (record === prior ? { ...record, runId } : record));
      return history([
        ...records,
        {
          version: 1,
          program,
          fingerprint,
          runId,
          updatedAt: new Date().toISOString(),
          stages: program.stages.map((stage) => ({
            id: stage.id,
            state: 'pending' as const,
            inspections: [],
          })),
        },
      ]);
    });
    try {
      const initial = (await this.inspect(program.id))!;
      if (initial.stages.some((stage) => stage.state === 'running' || stage.state === 'blocked'))
        throw new Error(
          'A stage has an unresolved or unverified outcome. Inspect it; it cannot be replayed automatically.',
        );
      const completed = initial.stages.filter((stage) => stage.state === 'verified');
      if (
        completed.length &&
        (!options.revalidate ||
          !(await options.revalidate(structuredClone(program), structuredClone(completed))))
      )
        throw new Error('Resume requires fresh source and completed-output validation.');
      for (const stage of program.stages) {
        options.signal?.throwIfAborted();
        const record = (await this.inspect(program.id))!;
        const checkpoint = record.stages.find((item) => item.id === stage.id)!;
        if (checkpoint.state === 'verified') {
          yield {
            type: 'durable-progress',
            programId: program.id,
            stageId: stage.id,
            state: 'resuming',
          };
          continue;
        }
        await this.store.update(this.scope, (raw) =>
          history(raw).map((record) => {
            if (record.program.id !== program.id) return record;
            const current = record.stages.find((item) => item.id === stage.id)!;
            if (
              record.runId !== runId ||
              current.state !== 'pending' ||
              record.stages.some((item) => item.state === 'running' || item.state === 'blocked') ||
              stage.dependsOn.some(
                (id) => record.stages.find((item) => item.id === id)?.state !== 'verified',
              )
            )
              throw new Error('Another pane owns this stage or a prerequisite is unresolved.');
            return {
              ...record,
              updatedAt: new Date().toISOString(),
              stages: record.stages.map((item) =>
                item.id === stage.id
                  ? {
                      id: item.id,
                      state: 'running' as const,
                      startedAt: new Date().toISOString(),
                      inspections: item.inspections,
                    }
                  : item,
              ),
            };
          }),
        );
        // Crashes/abandoned iterators leave the durable running claim. No lease expiry grants replay.
        let activeRequest: ActuationRequest | undefined;
        const checkpointEffect = async (rawRequest: ActuationRequest): Promise<void> => {
          const request = ActuationRequestSchema.parse(rawRequest);
          activeRequest = request;
          const effect = {
            changeId: request.changeId,
            kind: request.kind,
            fingerprint: await contentHash({
              surface: request.surface,
              kind: request.kind,
              params: request.params,
              preconditions: request.preconditions,
            }),
          };
          await this.store.update(this.scope, (raw) =>
            history(raw).map((record) => {
              if (record.program.id !== program.id) return record;
              const current = record.stages.find((item) => item.id === stage.id)!;
              if (record.runId !== runId || current.state !== 'running' || current.effect)
                throw new Error(
                  'The stage already has an effect claim or lost its execution owner.',
                );
              return {
                ...record,
                stages: record.stages.map((item) =>
                  item.id === stage.id ? { ...item, effect } : item,
                ),
              };
            }),
          );
        };
        const assertOwned = async (): Promise<void> => {
          const current = await this.inspect(program.id);
          if (
            current?.runId !== runId ||
            current.stages.find((item) => item.id === stage.id)?.state !== 'running'
          )
            throw new Error('This stage no longer owns its durable claim. Inspect its outcome.');
        };
        const receipts = (yield* execute(stage, checkpointEffect, assertOwned)).map((receipt) =>
          activeRequest
            ? validateActuationResult(activeRequest, receipt)
            : ActuationResultSchema.parse(receipt),
        );
        const expected = (await this.inspect(program.id))?.stages.find(
          (item) => item.id === stage.id,
        )?.effect;
        const correlated =
          receipts.length === 1 &&
          receipts[0]!.kind === expected?.kind &&
          receipts[0]!.changeId === expected?.changeId;
        const verified =
          correlated &&
          assessActuationResult(receipts[0]!) === 'verified' &&
          !receipts[0]!.recoveryPending;
        const cleanRefusal =
          correlated && ['failed', 'rejected'].includes(assessActuationResult(receipts[0]!));
        await this.store.update(this.scope, (raw) =>
          history(raw).map((record) => {
            if (record.program.id !== program.id) return record;
            if (
              record.runId !== runId ||
              record.stages.find((item) => item.id === stage.id)?.state !== 'running'
            )
              throw new Error(
                'The stage changed before its receipt checkpoint. Inspect the actual effect receipt.',
              );
            return {
              ...record,
              updatedAt: new Date().toISOString(),
              stages: record.stages.map((item) =>
                item.id !== stage.id
                  ? item
                  : {
                      ...item,
                      state: verified
                        ? ('verified' as const)
                        : cleanRefusal
                          ? ('pending' as const)
                          : ('blocked' as const),
                      ...(receipts.length === 1
                        ? { receipt: checkpointReceipt(receipts[0]!) }
                        : {}),
                    },
              ),
            };
          }),
        );
        yield {
          type: 'durable-progress',
          programId: program.id,
          stageId: stage.id,
          state: verified ? 'verified' : 'blocked',
        };
        if (!verified) return;
      }
    } finally {
      await this.store.update(this.scope, (raw) =>
        history(raw).map((record) => {
          if (
            record.program.id !== program.id ||
            record.runId !== runId ||
            record.stages.some((stage) => stage.state === 'running')
          )
            return record;
          const { runId: _runId, ...settled } = record;
          return settled;
        }),
      );
    }
  }
}
