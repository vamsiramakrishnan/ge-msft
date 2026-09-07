import { z } from 'zod';
import {
  ActuationRequestSchema,
  CellSnapshotSchema,
  asChangeId,
  cellsMatchRequest,
  hasFormulaErrors,
  gridForRequest,
  rangeForGrid,
  sourceVersion,
  assessActuationResult,
  validateActuationResult,
  type ActuationRequest,
  type ActuationResult,
  type CellSnapshot,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';

const RecordSchema = z.object({
  version: z.literal(1),
  id: z.string().max(256),
  owner: z.string().max(512),
  createdAt: z.string().datetime(),
  updatedAt: z.string().datetime(),
  state: z.enum([
    'applying',
    'applied',
    'not-applied',
    'uncertain',
    'conflict',
    'undone',
    'superseded',
  ]),
  request: ActuationRequestSchema,
  before: CellSnapshotSchema,
  afterHash: z.string().optional(),
  parentId: z.string().optional(),
  undo: z.boolean().optional(),
  message: z.string().max(1000).optional(),
});
type RecoveryRecord = z.infer<typeof RecordSchema>;
export interface RecoverySummary {
  id: string;
  createdAt: string;
  state: RecoveryRecord['state'];
  target: string;
  rows: number;
  columns: number;
  canUndo: boolean;
  canResume: boolean;
  canForget?: boolean;
  message?: string;
}
const HistorySchema = z.array(RecordSchema).max(32);
const now = (): string => new Date().toISOString();
const queues = new Map<string, Promise<unknown>>();
function isUnresolved(record: RecoveryRecord): boolean {
  return (
    record.state === 'applying' ||
    record.state === 'uncertain' ||
    // A verified write edited later has stale undo authority, but its original outcome is known.
    (record.state === 'conflict' && !record.afterHash)
  );
}

/** Worksheet IDs survive renames; host-captured locators supply canonical cell coordinates. */
function overlaps(left: CellSnapshot, right: CellSnapshot): boolean {
  if (left.surface !== right.surface || left.documentId !== right.documentId) return false;
  if (left.objectId && right.objectId && left.objectId !== right.objectId) return false;
  const sheet = (locator: string): string | undefined => {
    const bang = locator.lastIndexOf('!');
    if (bang < 0) return undefined;
    return locator
      .slice(0, bang)
      .replace(/^'(.*)'$/, '$1')
      .replace(/''/g, "'")
      .toLowerCase();
  };
  if (!left.objectId || !right.objectId) {
    const a = sheet(left.locator);
    const b = sheet(right.locator);
    if (a && b && a !== b) return false;
  }
  const bounds = (snapshot: CellSnapshot) => {
    const address = snapshot.locator.slice(snapshot.locator.lastIndexOf('!') + 1);
    const match = /^\$?([a-z]{1,3})\$?([1-9]\d*)(?::\$?([a-z]{1,3})\$?([1-9]\d*))?$/i.exec(address);
    if (!match) return undefined;
    const column = (value: string): number =>
      [...value.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    const x = column(match[1]!);
    const y = Number(match[2]);
    return {
      x,
      y,
      endX: Math.max(match[3] ? column(match[3]) : x, x + snapshot.values[0]!.length - 1),
      endY: Math.max(match[4] ? Number(match[4]) : y, y + snapshot.values.length - 1),
    };
  };
  const a = bounds(left);
  const b = bounds(right);
  // A malformed legacy locator cannot establish that a write on the same sheet is disjoint.
  return !a || !b || (a.x <= b.endX && b.x <= a.endX && a.y <= b.endY && b.y <= a.endY);
}

/** Durable effect receipts, never persisted approval authority. Only bounded cell writes are replayable. */
export class RecoveryCoordinator {
  private records: RecoveryRecord[] = [];
  private readonly prepared = new Map<string, CellSnapshot>();
  private readonly parents = new Map<string, { id: string; undo: boolean }>();
  constructor(
    private readonly bridge: DocBridge,
    private readonly owner: string,
  ) {}
  get durable(): boolean {
    return Boolean(this.bridge.recoveryStorage);
  }
  private async locked<T>(run: () => Promise<T>): Promise<T> {
    const key = `ge:recovery:${this.owner}`;
    const previous = queues.get(key) ?? Promise.resolve();
    const promise = previous
      .catch(() => {})
      .then(async (): Promise<T> => {
        // Serializes same-origin panes. Office settings are not a distributed compare-and-swap;
        // cross-device changes are handled by source/readback checks, never automatic replay.
        return await (typeof navigator !== 'undefined' && navigator.locks
          ? navigator.locks.request(key, run)
          : run());
      });
    queues.set(key, promise);
    try {
      return await promise;
    } finally {
      if (queues.get(key) === promise) queues.delete(key);
    }
  }
  private async load(): Promise<void> {
    if (!this.bridge.recoveryStorage) return;
    const records = HistorySchema.parse(await this.bridge.recoveryStorage.load(this.owner));
    if (records.some((r) => r.owner !== this.owner))
      throw new Error('Recovery history belongs to a different identity.');
    this.records = records;
  }
  private async save(): Promise<void> {
    await this.bridge.recoveryStorage?.save(this.owner, this.records);
  }
  private assertNoUnresolvedOverlap(before: CellSnapshot): void {
    const unresolved = this.records.find(
      (record) => isUnresolved(record) && overlaps(record.before, before),
    );
    if (unresolved)
      throw new Error(
        `An unresolved write overlaps ${before.locator}. Inspect recovery history before writing to these cells again.`,
      );
  }

  /** Capture before approval. An absolute, versioned destination replaces the live selection. */
  async prepare(request: ActuationRequest): Promise<ActuationRequest> {
    if (request.kind !== 'write-cells' || !this.bridge.captureCells) return request;
    const prepared = this.prepared.get(request.changeId);
    if (prepared) {
      await this.locked(async () => {
        await this.load();
        this.assertNoUnresolvedOverlap(prepared);
      });
      return request;
    }
    const values = gridForRequest(request);
    const address = rangeForGrid(
      request.params.target?.range ?? '',
      values.length,
      values[0]?.length ?? 0,
    );
    const before = await this.bridge.captureCells(address);
    const conditions = request.preconditions ?? [];
    for (const expected of conditions) {
      const current =
        expected.locator === before.locator
          ? before
          : await this.bridge.captureCells(expected.locator);
      if (
        current.hash !== expected.hash ||
        current.documentId !== expected.documentId ||
        current.surface !== expected.surface
      )
        throw new Error(
          'A source or destination changed. Capture it again before reviewing this write.',
        );
    }
    await this.locked(async () => {
      await this.load();
      this.assertNoUnresolvedOverlap(before);
      this.prepared.set(request.changeId, before);
    });
    return ActuationRequestSchema.parse({
      ...request,
      params: { ...request.params, target: { ...request.params.target, range: before.locator } },
      preconditions: [
        ...conditions.filter((c) => c.locator !== before.locator),
        sourceVersion(before),
      ],
    });
  }
  clearPrepared(): void {
    this.prepared.clear();
    this.parents.clear();
  }

  /** A failed pre-write checkpoint prevents mutation; a failed post-write checkpoint preserves success. */
  async execute(
    request: ActuationRequest,
    actuate: () => Promise<ActuationResult>,
  ): Promise<ActuationResult> {
    const before = this.prepared.get(request.changeId);
    if (request.kind !== 'write-cells' || !before)
      return validateActuationResult(request, await actuate());
    return this.locked(async () => {
      await this.load();
      if (this.records.some((r) => r.id === request.changeId))
        throw new Error(
          'This effect already has a receipt. Review recovery history before retrying.',
        );
      if (this.records.length >= 32)
        throw new Error('Recovery history is full. Remove a resolved receipt before writing.');
      this.assertNoUnresolvedOverlap(before);
      const parent = this.parents.get(request.changeId);
      if (parent) {
        const original = this.records.find((r) => r.id === parent.id);
        if (
          !original ||
          (parent.undo
            ? original.state !== 'applied' || !original.afterHash
            : original.state !== 'not-applied')
        )
          throw new Error('The recovery receipt changed. Inspect it again before continuing.');
      }
      const record: RecoveryRecord = {
        version: 1,
        id: request.changeId,
        owner: this.owner,
        createdAt: now(),
        updatedAt: now(),
        state: 'applying',
        request: structuredClone(request),
        before,
        ...(parent ? { parentId: parent.id, undo: parent.undo } : {}),
      };
      this.records.push(record);
      try {
        await this.save();
      } catch (error) {
        this.records.pop();
        throw error;
      }
      let result: ActuationResult;
      try {
        result = validateActuationResult(request, await actuate());
      } catch {
        // Office can reject a batch after a write landed. Do not turn this into retry authority.
        record.state = 'uncertain';
        record.message =
          'Host execution was interrupted. Inspect the cells before deciding what to do.';
        try {
          await this.save();
        } catch {
          /* The durable applying intent remains recoverable. */
        }
        return {
          ok: false,
          kind: request.kind,
          changeId: request.changeId,
          recoveryPending: true,
          error: { code: 'outcome_unknown', message: record.message },
        };
      }
      const outcome = assessActuationResult(result);
      record.state =
        outcome === 'verified'
          ? 'applied'
          : outcome === 'uncertain' || outcome === 'unverified'
            ? 'uncertain'
            : 'not-applied';
      record.afterHash =
        result.verification?.status === 'verified' ? result.verification.afterHash : undefined;
      record.message = result.verification?.message ?? result.error?.message;
      record.updatedAt = now();
      if (record.state === 'applied' && parent) {
        const original = this.records.find((r) => r.id === parent.id);
        if (original) {
          original.state = parent.undo ? 'undone' : 'superseded';
          original.updatedAt = now();
        }
      }
      try {
        await this.save();
        return result;
      } catch {
        return { ...result, recoveryPending: true };
      }
    });
  }
  async inspect(): Promise<RecoverySummary[]> {
    return this.locked(async () => {
      await this.load();
      if (!this.bridge.captureCells) return [];
      for (const record of this.records) {
        if (record.state === 'undone' || record.state === 'superseded') continue;
        try {
          const current = await this.bridge.captureCells(record.before.locator);
          if (
            current.documentId !== record.before.documentId ||
            current.objectId !== record.before.objectId
          ) {
            record.state = 'conflict';
            record.message = 'The original document or worksheet is no longer active.';
          } else if (record.afterHash && current.hash === record.afterHash) {
            record.state = 'applied';
          } else if (current.hash === record.before.hash) {
            record.state = 'not-applied';
          } else if (
            !record.afterHash &&
            cellsMatchRequest(current, record.request) &&
            !hasFormulaErrors(current)
          ) {
            record.state = 'applied';
            record.afterHash = current.hash;
          } else {
            record.state = 'conflict';
            record.message = record.afterHash
              ? 'This verified write was changed later. Undo is no longer safe; review a new write or remove this receipt.'
              : 'The cells differ from both the reviewed input and output. Review them manually.';
          }
        } catch {
          record.state = 'uncertain';
          record.message =
            'The destination could not be read. Open the original worksheet and try again.';
        }
      }
      await this.save();
      return this.list();
    });
  }
  list(): RecoverySummary[] {
    return this.records
      .slice()
      .reverse()
      .map((r) => ({
        id: r.id,
        createdAt: r.createdAt,
        state: r.state,
        target: r.before.locator,
        rows: r.before.values.length,
        columns: r.before.values[0]!.length,
        canUndo: r.state === 'applied' && Boolean(r.afterHash),
        canResume: r.state === 'not-applied',
        canForget: !isUnresolved(r),
        ...(r.message ? { message: r.message } : {}),
      }));
  }
  async request(id: string, undo: boolean): Promise<ActuationRequest> {
    await this.inspect();
    const record = this.records.find((r) => r.id === id);
    if (
      !record ||
      (undo ? record.state !== 'applied' || !record.afterHash : record.state !== 'not-applied')
    )
      throw new Error(
        'This receipt cannot be replayed in its current state. Refresh recovery history.',
      );
    const changeId = asChangeId(`recovery:${crypto.randomUUID()}`);
    const request: ActuationRequest = undo
      ? {
          changeId,
          surface: 'excel',
          kind: 'write-cells',
          params: {
            target: { range: record.before.locator },
            cellValues: record.before.values,
            ...(record.before.formulas ? { cellFormulas: record.before.formulas } : {}),
          },
          preconditions: [{ ...sourceVersion(record.before), hash: record.afterHash! }],
        }
      : { ...structuredClone(record.request), changeId };
    this.parents.set(changeId, { id, undo });
    return request;
  }
  async forget(id: string): Promise<void> {
    await this.locked(async () => {
      await this.load();
      const entry = this.records.find((r) => r.id === id);
      if (entry && isUnresolved(entry)) throw new Error('Unresolved receipts cannot be discarded.');
      this.records = this.records.filter((r) => r.id !== id);
      await this.save();
    });
  }
}
