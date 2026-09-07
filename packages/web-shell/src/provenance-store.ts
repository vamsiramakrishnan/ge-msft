import type { ActuationResult, ChangeId, ProvenancePayload } from '@ge/contracts';

/**
 * The panel's view-model of what the agent has changed — the "traceable + reversible" surface.
 * Durable provenance lives in the host's own metadata (the bridge writes it); this is the
 * in-session record the UI lists so the user can see who/what/why and undo.
 */
export interface ChangeRecord {
  changeId: ChangeId;
  kind: ActuationResult['kind'];
  ok: boolean;
  location?: string;
  degraded?: boolean;
  error?: { code: string; message: string };
  /** Preserve verification and recovery truth instead of collapsing a receipt to `ok`. */
  verification?: ActuationResult['verification'];
  recoveryPending?: ActuationResult['recoveryPending'];
  provenance?: ProvenancePayload;
  /** Preserve host-provided inverse receipts for a future explicit, conflict-checked undo action. */
  inverse?: ActuationResult['inverse'];
  /** When the panel recorded it (client clock). */
  at: string;
}

/**
 * A `share` (estate write) the panel has recorded — the same "who/what/why" audit surface as
 * `ChangeRecord`, but for `/shared` writes, which carry no `changeId`/`ActuationKind` (they never
 * reach `bridge.actuate()`). Kept as its own list rather than folded into `ChangeRecord` so no
 * existing `ChangeRecord` consumer has to widen its assumptions about what a "change" looks like.
 */
export interface ShareRecord {
  name: string;
  bytes: number;
  sourceLabel: string;
  truncated: boolean;
  provenance?: ProvenancePayload;
  /** When the panel recorded it (client clock). */
  at: string;
}

export class ProvenanceStore {
  private readonly records = new Map<ChangeId, ChangeRecord>();
  private readonly shares: ShareRecord[] = [];

  constructor(private readonly now: () => string = () => new Date().toISOString()) {}

  /** Record the outcome of an actuation, stamped with the turn's provenance when available. */
  record(result: ActuationResult, provenance?: ProvenancePayload): ChangeRecord {
    const rec: ChangeRecord = {
      changeId: result.changeId,
      kind: result.kind,
      ok: result.ok,
      ...(result.location ? { location: result.location } : {}),
      ...(result.degraded ? { degraded: true } : {}),
      ...(result.error ? { error: { ...result.error } } : {}),
      ...(result.verification ? { verification: structuredClone(result.verification) } : {}),
      ...(result.recoveryPending !== undefined ? { recoveryPending: result.recoveryPending } : {}),
      ...(provenance ? { provenance: structuredClone(provenance) } : {}),
      ...(result.inverse ? { inverse: structuredClone(result.inverse) } : {}),
      at: this.now(),
    };
    this.records.set(rec.changeId, rec);
    return structuredClone(rec);
  }

  /** Record a successful `share` — the `/shared` analog of {@link record}, for the audit surface. */
  recordShare(
    input: { name: string; bytes: number; sourceLabel: string; truncated: boolean },
    provenance?: ProvenancePayload,
  ): ShareRecord {
    const rec: ShareRecord = {
      ...input,
      ...(provenance ? { provenance: structuredClone(provenance) } : {}),
      at: this.now(),
    };
    this.shares.push(rec);
    return structuredClone(rec);
  }

  get(changeId: ChangeId): ChangeRecord | undefined {
    const record = this.records.get(changeId);
    return record ? structuredClone(record) : undefined;
  }

  list(): ChangeRecord[] {
    return structuredClone([...this.records.values()]);
  }

  listShares(): ShareRecord[] {
    return structuredClone(this.shares);
  }

  get size(): number {
    return this.records.size;
  }
}
