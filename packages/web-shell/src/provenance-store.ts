import type { ActuationResult, ProvenancePayload } from '@ge/contracts';

/**
 * The panel's view-model of what the agent has changed — the "traceable + reversible" surface.
 * Durable provenance lives in the host's own metadata (the bridge writes it); this is the
 * in-session record the UI lists so the user can see who/what/why and undo.
 */
export interface ChangeRecord {
  changeId: string;
  kind: ActuationResult['kind'];
  ok: boolean;
  location?: string;
  degraded?: boolean;
  error?: { code: string; message: string };
  provenance?: ProvenancePayload;
  /** When the panel recorded it (client clock). */
  at: string;
}

export class ProvenanceStore {
  private readonly records = new Map<string, ChangeRecord>();

  constructor(private readonly now: () => string = () => new Date().toISOString()) {}

  /** Record the outcome of an actuation, stamped with the turn's provenance when available. */
  record(result: ActuationResult, provenance?: ProvenancePayload): ChangeRecord {
    const rec: ChangeRecord = {
      changeId: result.changeId,
      kind: result.kind,
      ok: result.ok,
      ...(result.location ? { location: result.location } : {}),
      ...(result.degraded ? { degraded: true } : {}),
      ...(result.error ? { error: result.error } : {}),
      ...(provenance ? { provenance } : {}),
      at: this.now(),
    };
    this.records.set(rec.changeId, rec);
    return rec;
  }

  get(changeId: string): ChangeRecord | undefined {
    return this.records.get(changeId);
  }

  list(): ChangeRecord[] {
    return [...this.records.values()];
  }

  get size(): number {
    return this.records.size;
  }
}
