/**
 * The request in progress, kept in this browser so a task cut short by closing or reloading the pane
 * can be finished later. Written when a change task starts, updated as each change lands, and removed
 * when the task ends in the pane (done, failed or cancelled), so a record that survives a reopen is an
 * interrupted task.
 *
 * Stored: the request the user typed (capped), short labels of the changes that landed (these can
 * name a sheet, a slide title, or the first line of text written into a slide shape) and a hash of
 * the document. No tokens. Per signed-in account (a hash,
 * as saved diagnostics), offered only in the same document, ignored after 24 hours.
 */
export interface InProgress {
  /** Hash of the document the task ran in: a record is offered only in that document. */
  doc: string;
  request: string;
  /** Plain labels of the changes that already landed ("Cells Sheet1!F2:F11"). */
  changes: string[];
  startedAt: string;
}

type KeyValueStorage = Pick<Storage, 'getItem' | 'setItem' | 'removeItem'>;

const KEY = 'ge.in-progress.v1';
const MAX_AGE_MS = 24 * 60 * 60 * 1000;
const MAX_REQUEST = 1000;
const MAX_CHANGES = 50;

export class ResumeStore {
  private readonly key: string;

  /** Without a signed-in scope nothing is stored. */
  constructor(
    private storage: KeyValueStorage | undefined,
    scope: string | undefined,
    private readonly now: () => number = Date.now,
  ) {
    this.key = `${KEY}:${scope ?? ''}`;
    if (!scope) this.storage = undefined;
  }

  /** The record for this document, if any (`doc` is the hash from {@link documentHash}). */
  read(doc?: string): InProgress | undefined {
    try {
      const raw = this.storage?.getItem(this.key);
      if (!raw) return undefined;
      const r = JSON.parse(raw) as Partial<InProgress>;
      if (doc !== undefined && r.doc !== doc) return undefined;
      if (
        typeof r.doc !== 'string' ||
        typeof r.request !== 'string' ||
        !Array.isArray(r.changes) ||
        typeof r.startedAt !== 'string' ||
        this.now() - Date.parse(r.startedAt) > MAX_AGE_MS
      ) {
        this.clear();
        return undefined;
      }
      return {
        doc: r.doc,
        request: r.request,
        changes: r.changes.filter((c): c is string => typeof c === 'string'),
        startedAt: r.startedAt,
      };
    } catch {
      return undefined;
    }
  }

  start(request: string, doc: string): void {
    this.write({
      doc,
      request: request.trim().slice(0, MAX_REQUEST),
      changes: [],
      startedAt: new Date(this.now()).toISOString(),
    });
  }

  landed(label: string): void {
    const current = this.read();
    if (current)
      this.write({ ...current, changes: [...current.changes, label].slice(0, MAX_CHANGES) });
  }

  clear(): void {
    try {
      this.storage?.removeItem(this.key);
    } catch {
      /* Storage blocked: nothing was stored. */
    }
  }

  private write(value: InProgress): void {
    try {
      this.storage?.setItem(this.key, JSON.stringify(value));
    } catch {
      /* Best effort: resuming is a convenience, never a reason to fail a task. */
    }
  }
}

/** A short SHA-256 of the document key, so the record never holds the document's URL. */
export async function documentHash(key: string): Promise<string | undefined> {
  try {
    const digest = await crypto.subtle.digest(
      'SHA-256',
      new TextEncoder().encode(`ge-in-progress:${key}`),
    );
    return [...new Uint8Array(digest)]
      .slice(0, 12)
      .map((b) => b.toString(16).padStart(2, '0'))
      .join('');
  } catch {
    // No Web Crypto (some hosts and tests): a short FNV-1a hash still keeps the URL out of storage.
    let h = 0x811c9dc5;
    for (const ch of `ge-in-progress:${key}`) {
      h ^= ch.codePointAt(0)!;
      h = Math.imul(h, 0x01000193) >>> 0;
    }
    return `f${h.toString(16)}`;
  }
}
