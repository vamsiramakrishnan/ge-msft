/**
 * Saved diagnostics: `{ question, diagnostic }` per request, kept in this browser so a user can still
 * send support the diagnostics after the pane or Office was closed (the chat itself is gone then).
 * A product-owner-approved exception to "no request text in browser storage" (ADR-0012 amendment).
 *
 * What is stored: the text the user typed in the composer (capped) and the content-free diagnostics
 * snapshot (routes, step kinds, redacted errors, run outcomes). Never document text, model output,
 * history loaded from the server, or tokens. Written only after the user sent something, under a key
 * scoped to the signed-in account (a hash, never the raw id). Bounded at rest: the newest 20 entries,
 * none older than 7 days, about 512 K characters in total; anything past that is deleted, oldest first.
 */
export interface SavedDiagnostic {
  /** One typed request: runs it starts later (e.g. after a plan is confirmed) update this entry. */
  id: string;
  /** ISO time of the latest update. */
  at: string;
  question: string;
  diagnostic: unknown;
}

type KeyValueStorage = Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>;

const KEY = 'ge.saved-diagnostics.v1';
const MAX_ENTRIES = 20;
const MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
const MAX_CHARS = 512 * 1024;
const MAX_QUESTION = 1000;

export class SavedDiagnostics {
  private readonly key: string;

  /** `scope`: the signed-in account's hash ({@link diagnosticsScope}); without one nothing is stored. */
  constructor(
    private storage: KeyValueStorage | undefined,
    scope: string | undefined,
    private readonly now: () => number = Date.now,
  ) {
    this.key = `${KEY}:${scope ?? ''}`;
    if (!scope) this.storage = undefined;
  }

  /** Newest first. Expired entries are deleted from storage, not only hidden. */
  list(): SavedDiagnostic[] {
    const stored = this.read();
    const kept = this.prune(stored);
    if (kept.length !== stored.length) this.write(kept);
    return kept;
  }

  save(entry: { id: string; question: string; diagnostic: unknown }): void {
    const question = entry.question.trim().slice(0, MAX_QUESTION);
    if (!question) return;
    const saved: SavedDiagnostic = {
      id: entry.id,
      at: new Date(this.now()).toISOString(),
      question,
      diagnostic: entry.diagnostic,
    };
    this.write([saved, ...this.read().filter((e) => e.id !== entry.id)]);
  }

  clear(): void {
    try {
      this.storage?.removeItem(this.key);
    } catch {
      /* Storage blocked (private mode, policy): nothing was saved. */
    }
  }

  private read(): SavedDiagnostic[] {
    try {
      const raw = this.storage?.getItem(this.key);
      const parsed: unknown = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed.filter(isEntry) : [];
    } catch {
      return [];
    }
  }

  private write(entries: SavedDiagnostic[]): void {
    let kept = this.prune(entries);
    let json = JSON.stringify(kept);
    while (kept.length > 1 && json.length > MAX_CHARS) {
      kept = kept.slice(0, -1);
      json = JSON.stringify(kept);
    }
    try {
      if (kept.length === 0) this.storage?.removeItem(this.key);
      else if (json.length <= MAX_CHARS) this.storage?.setItem(this.key, json);
    } catch {
      /* Quota or policy: saving diagnostics is best effort and never blocks the pane. */
    }
  }

  private prune(entries: SavedDiagnostic[]): SavedDiagnostic[] {
    const cutoff = this.now() - MAX_AGE_MS;
    return entries
      .filter((e) => Date.parse(e.at) >= cutoff)
      .sort((a, b) => b.at.localeCompare(a.at))
      .slice(0, MAX_ENTRIES);
  }
}

function isEntry(value: unknown): value is SavedDiagnostic {
  const e = value as Partial<SavedDiagnostic> | null;
  return (
    typeof e === 'object' &&
    e !== null &&
    typeof e.id === 'string' &&
    typeof e.at === 'string' &&
    typeof e.question === 'string'
  );
}

/** A short SHA-256 of the account id, so the storage key never holds the raw id or email. */
export async function diagnosticsScope(accountId: string): Promise<string | undefined> {
  try {
    const bytes = new TextEncoder().encode(`ge-saved-diagnostics:${accountId}`);
    const digest = await crypto.subtle.digest('SHA-256', bytes);
    return [...new Uint8Array(digest)]
      .slice(0, 12)
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('');
  } catch {
    return undefined;
  }
}

/** `window.localStorage`, or undefined where the host blocks it. */
export function browserStorage(): KeyValueStorage | undefined {
  try {
    return typeof window === 'undefined' ? undefined : window.localStorage;
  } catch {
    return undefined;
  }
}
