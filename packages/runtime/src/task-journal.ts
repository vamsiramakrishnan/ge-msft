import { z } from 'zod';
import {
  ActuationRequestSchema,
  ActuationResultSchema,
  assessActuationResult,
  contentHash,
  unknownActuationResult,
  validateActuationResult,
  type ActuationRequest,
  type ActuationResult,
} from '@ge/contracts';

const EntrySchema = z
  .object({
    version: z.literal(1),
    id: z.string().min(1).max(256),
    taskId: z.string().min(1).max(256),
    fingerprint: z.string().regex(/^sha256:[a-f0-9]{64}$/),
    createdAt: z.string().datetime(),
    updatedAt: z.string().datetime(),
    state: z.enum(['dispatching', 'applied', 'not-applied', 'uncertain', 'acknowledged']),
    payloadOmitted: z.boolean().default(false),
    request: ActuationRequestSchema,
    receipt: ActuationResultSchema.optional(),
  })
  .strict();
export type TaskJournalEntry = z.infer<typeof EntrySchema>;
const MAX_ENTRIES = 1024;
const MAX_BYTES = 8 * 1024 * 1024;
const HistorySchema = z.array(EntrySchema).max(MAX_ENTRIES);
function validateHistory(raw: unknown): TaskJournalEntry[] {
  const entries = HistorySchema.parse(raw === undefined ? [] : raw);
  if (new TextEncoder().encode(JSON.stringify(entries)).byteLength > MAX_BYTES)
    throw new Error(
      'Task journal quota exceeded. Archive inspected completed work before continuing.',
    );
  return entries;
}
export interface TaskJournalStore {
  readonly durable: boolean;
  read(scope: string): Promise<unknown>;
  /** A synchronous mutation executed in one storage transaction. No host work is allowed inside it. */
  update(scope: string, change: (history: unknown) => TaskJournalEntry[]): Promise<void>;
}
export class MemoryTaskJournalStore implements TaskJournalStore {
  readonly durable = false;
  private readonly histories = new Map<string, TaskJournalEntry[]>();
  async read(scope: string): Promise<unknown> {
    return structuredClone(this.histories.get(scope) ?? []);
  }
  async update(scope: string, change: (history: unknown) => TaskJournalEntry[]): Promise<void> {
    this.histories.set(
      scope,
      structuredClone(validateHistory(change(structuredClone(this.histories.get(scope) ?? [])))),
    );
  }
}
/** IndexedDB serializes read/write transactions across panes, including the pre-dispatch claim. */
export class IndexedDbTaskJournalStore implements TaskJournalStore {
  readonly durable = true;
  private readonly database: Promise<IDBDatabase>;
  constructor(factory: IDBFactory = indexedDB) {
    this.database = new Promise((resolve, reject) => {
      const request = factory.open('ge-msft-task-journal', 1);
      request.onupgradeneeded = () => request.result.createObjectStore('histories');
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onerror = () => reject(new Error('Task journal could not open.'));
      request.onblocked = () =>
        reject(new Error('Close an older add-in pane to upgrade task storage.'));
    });
  }
  async read(scope: string): Promise<unknown> {
    const database = await this.database;
    return new Promise((resolve, reject) => {
      const transaction = database.transaction('histories', 'readonly');
      const request = transaction.objectStore('histories').get(scope);
      transaction.oncomplete = () => resolve(request.result === undefined ? [] : request.result);
      transaction.onabort = transaction.onerror = () =>
        reject(new Error('Task journal read failed.'));
    });
  }
  async update(scope: string, change: (history: unknown) => TaskJournalEntry[]): Promise<void> {
    const database = await this.database;
    return new Promise((resolve, reject) => {
      const transaction = database.transaction('histories', 'readwrite');
      const store = transaction.objectStore('histories');
      const request = store.get(scope);
      let failure: unknown;
      request.onsuccess = () => {
        try {
          store.put(
            validateHistory(change(request.result === undefined ? [] : request.result)),
            scope,
          );
        } catch (error) {
          failure = error;
          transaction.abort();
        }
      };
      transaction.oncomplete = () => resolve();
      transaction.onabort = transaction.onerror = () =>
        reject(failure ?? new Error('Task journal write failed.'));
    });
  }
}
export interface TaskJournalSummary {
  id: string;
  taskId: string;
  kind: ActuationRequest['kind'];
  state: TaskJournalEntry['state'];
  updatedAt: string;
}
function mergeBaseKey(entry: TaskJournalEntry): string | undefined {
  const target = entry.request.params.target;
  if (entry.state !== 'applied' || entry.receipt?.verification?.status !== 'verified')
    return undefined;
  if (entry.request.kind === 'set-shape-text' && target?.slideId && target.shapeId)
    return `powerpoint:${target.slideId}:${target.shapeId}`;
  if (entry.request.kind === 'fill-content-control' && target?.contentControlId)
    return `word:${target.contentControlId}`;
  return undefined;
}
/** Full payloads and inverse bodies belong to artifact/recovery stores, never the dispatch log. */
function journalIntent(request: ActuationRequest): ActuationRequest {
  const managed = ['set-shape-text', 'fill-content-control'].includes(request.kind);
  const text =
    managed &&
    request.params.text !== undefined &&
    new TextEncoder().encode(request.params.text).byteLength <= 64 * 1024
      ? request.params.text
      : undefined;
  return ActuationRequestSchema.parse({
    changeId: request.changeId,
    kind: request.kind,
    surface: request.surface,
    params: {
      ...(request.params.target ? { target: request.params.target } : {}),
      ...(text !== undefined ? { text } : {}),
    },
    ...(request.preconditions ? { preconditions: request.preconditions } : {}),
  });
}
function journalReceipt(receipt: ActuationResult): ActuationResult {
  const { inverse: _inverse, ...summary } = receipt;
  return {
    ...summary,
    ...(receipt.location && receipt.location.length > 2048 ? { location: undefined } : {}),
    ...(receipt.error
      ? {
          error: {
            ...receipt.error,
            message: 'Inspect the recorded host outcome before retrying.',
          },
        }
      : {}),
  };
}
function appendBounded(history: TaskJournalEntry[], entry: TaskJournalEntry): TaskJournalEntry[] {
  const bases = new Set<string>();
  const protectedIds = new Set([entry.id]);
  for (const prior of [...history].reverse()) {
    const key = mergeBaseKey(prior);
    if (key && !bases.has(key)) {
      bases.add(key);
      protectedIds.add(prior.id);
    }
    if (['dispatching', 'uncertain'].includes(prior.state)) protectedIds.add(prior.id);
  }
  const next = [...history, entry];
  while (
    next.length > MAX_ENTRIES ||
    new TextEncoder().encode(JSON.stringify(next)).byteLength > MAX_BYTES
  ) {
    const remove = next.findIndex((item) => !protectedIds.has(item.id));
    if (remove < 0)
      throw new Error(
        'Task journal merge-base quota exceeded. Archive inspected outputs explicitly before continuing.',
      );
    next.splice(remove, 1);
  }
  return validateHistory(next);
}
/** A journal persists intent and outcomes, never approvals. Only the existing approved executor calls execute. */
export class TaskJournal {
  private entries: TaskJournalEntry[] = [];
  constructor(
    private readonly store: TaskJournalStore,
    private readonly scope: string,
  ) {
    if (!/^sha256:[a-f0-9]{64}$/.test(scope))
      throw new Error('Journal scope must hash principal and document identity.');
  }
  /** Hashed principal/document scope; contains no bearer or raw document identity. */
  get scopeKey(): string {
    return this.scope;
  }
  get durable(): boolean {
    return this.store.durable;
  }
  async inspect(): Promise<TaskJournalSummary[]> {
    this.entries = validateHistory(await this.store.read(this.scope));
    return this.list();
  }
  list(): TaskJournalSummary[] {
    return this.entries.map((entry) => ({
      id: entry.id,
      taskId: entry.taskId,
      kind: entry.request.kind,
      state: entry.state,
      updatedAt: entry.updatedAt,
    }));
  }
  /** A previously verified output is a merge base, never evidence of the current host contents. */
  lastGeneratedText(
    surface: ActuationRequest['surface'],
    objectId: string,
    parentObjectId?: string,
  ): string | undefined {
    const entry = [...this.entries]
      .reverse()
      .find(
        (entry) =>
          entry.state === 'applied' &&
          entry.receipt?.verification?.status === 'verified' &&
          entry.request.surface === surface &&
          (surface === 'powerpoint'
            ? entry.request.kind === 'set-shape-text' &&
              entry.request.params.target?.shapeId === objectId &&
              entry.request.params.target?.slideId === parentObjectId
            : surface === 'word' &&
              entry.request.kind === 'fill-content-control' &&
              entry.request.params.target?.contentControlId === objectId),
      );
    return entry?.request.params.text;
  }
  async execute(
    taskId: string,
    input: ActuationRequest,
    dispatch: () => Promise<ActuationResult>,
  ): Promise<ActuationResult> {
    const request = ActuationRequestSchema.parse(input);
    const fingerprint = await contentHash({
      surface: request.surface,
      kind: request.kind,
      params: request.params,
      preconditions: request.preconditions,
    });
    const now = new Date().toISOString();
    const entry: TaskJournalEntry = {
      version: 1,
      id: request.changeId,
      taskId,
      request: journalIntent(request),
      payloadOmitted: true,
      fingerprint,
      createdAt: now,
      updatedAt: now,
      state: 'dispatching',
    };
    await this.store.update(this.scope, (raw) => {
      const history = validateHistory(raw);
      if (history.some((prior) => ['dispatching', 'uncertain'].includes(prior.state)))
        throw new Error(
          'An earlier Office change has an unresolved outcome. Inspect and reconcile it before another write.',
        );
      if (history.some((prior) => prior.id === request.changeId))
        throw new Error(
          'This change ID is already journaled. Inspect its receipt before retrying.',
        );
      // Retain unresolved work indefinitely; prune only known, older outcomes when capacity is needed.
      return appendBounded(history, entry);
    });
    this.entries = validateHistory(await this.store.read(this.scope));
    let receipt: ActuationResult;
    try {
      receipt = validateActuationResult(request, await dispatch());
    } catch {
      receipt = unknownActuationResult(
        request,
        'Office execution was interrupted. Inspect the document before retrying.',
      );
    }
    const assessment = assessActuationResult(receipt);
    try {
      await this.store.update(this.scope, (raw) =>
        validateHistory(raw).map((prior) =>
          prior.id === entry.id
            ? {
                ...prior,
                updatedAt: new Date().toISOString(),
                receipt: journalReceipt(receipt),
                state:
                  assessment === 'uncertain' ? 'uncertain' : receipt.ok ? 'applied' : 'not-applied',
              }
            : prior,
        ),
      );
      this.entries = validateHistory(await this.store.read(this.scope));
      return receipt;
    } catch {
      // Keep the pre-dispatch claim. A failed checkpoint must never invite an automatic resend.
      return { ...receipt, recoveryPending: true };
    }
  }
  /** Explicit operator acknowledgement after inspection; does not replay a request or restore an approval. */
  async acknowledge(id: string): Promise<void> {
    await this.store.update(this.scope, (raw) =>
      validateHistory(raw).map((entry) => {
        if (entry.id !== id) return entry;
        if (
          entry.state === 'dispatching' &&
          Date.now() - Date.parse(entry.updatedAt) < 5 * 60 * 1000
        )
          throw new Error(
            'This change may still be in flight in another pane. Wait five minutes before acknowledging inspection.',
          );
        return { ...entry, state: 'acknowledged', updatedAt: new Date().toISOString() };
      }),
    );
    await this.inspect();
  }
}
