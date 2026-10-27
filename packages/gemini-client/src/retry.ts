/**
 * Resilience helpers for the idempotent read paths (search, autocomplete, grounding,
 * rank, the WIF STS exchange). Retries only *transient* failures with exponential
 * backoff and full jitter; never retries deterministic 4xx (except 429) or any
 * non-idempotent operation — that is the caller's responsibility.
 *
 * The sleep and RNG are injectable so backoff is deterministic under test. No Google
 * credential or token ever flows through here; errors are surfaced unchanged.
 */

export interface RetryOptions {
  /** Total attempts including the first (>= 1). Default 3. */
  maxAttempts?: number;
  /** Base backoff in ms; the exponential schedule is base * 2**attempt. Default 200. */
  baseMs?: number;
  /** Upper bound on a single backoff window in ms. Default 5000. */
  capMs?: number;
  /** Decides whether a thrown error / rejected value is worth retrying. */
  isRetriable?: (err: unknown) => boolean;
  /** Injected sleep (ms). Defaults to a real timer. */
  sleep?: (ms: number) => Promise<void>;
  /** Injected RNG in [0, 1). Defaults to Math.random. */
  random?: () => number;
  /** Optional circuit breaker consulted before each call and updated after it. */
  breaker?: CircuitBreaker;
}

const DEFAULT_MAX_ATTEMPTS = 3;
const DEFAULT_BASE_MS = 200;
const DEFAULT_CAP_MS = 5000;
/** The longest a `Retry-After` header may hold a retry; longer values are treated as this. */
export const MAX_RETRY_AFTER_MS = 30_000;

/**
 * Parse a `Retry-After` header (delay-seconds or an HTTP date) into milliseconds from `now`,
 * capped at {@link MAX_RETRY_AFTER_MS}. Undefined when absent or unparseable.
 */
export function parseRetryAfter(
  header: string | null | undefined,
  now: () => number = Date.now,
): number | undefined {
  if (!header) return undefined;
  const text = header.trim();
  if (text === '') return undefined;
  let ms: number;
  if (/^\d+$/.test(text)) ms = Number(text) * 1000;
  else {
    const at = Date.parse(text);
    if (Number.isNaN(at)) return undefined;
    ms = at - now();
  }
  if (!Number.isFinite(ms) || ms < 0) ms = 0;
  return Math.min(ms, MAX_RETRY_AFTER_MS);
}

function defaultSleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Full-jitter backoff: `random() * min(cap, base * 2**attempt)`. `attempt` is the
 * zero-based index of the *failed* attempt about to be retried, so the first retry
 * draws from [0, base), the second from [0, 2*base), etc., each clamped to the cap.
 */
export function backoffDelay(
  attempt: number,
  baseMs: number,
  capMs: number,
  random: () => number,
): number {
  const exp = baseMs * 2 ** attempt;
  const window = Math.min(capMs, exp);
  return Math.floor(random() * window);
}

/**
 * Run `fn` with exponential backoff + full jitter, retrying only when `isRetriable`
 * says so and attempts remain. The last error is rethrown unchanged once attempts are
 * exhausted, so callers see the real cause (e.g. the original HTTP error string).
 */
export async function withRetry<T>(fn: () => Promise<T>, opts: RetryOptions = {}): Promise<T> {
  const maxAttempts = Math.max(1, opts.maxAttempts ?? DEFAULT_MAX_ATTEMPTS);
  const baseMs = opts.baseMs ?? DEFAULT_BASE_MS;
  const capMs = opts.capMs ?? DEFAULT_CAP_MS;
  const isRetriable = opts.isRetriable ?? defaultIsRetriable;
  const sleep = opts.sleep ?? defaultSleep;
  const random = opts.random ?? Math.random;
  const breaker = opts.breaker;

  let lastErr: unknown;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    breaker?.assertClosed();
    try {
      const result = await fn();
      breaker?.recordSuccess();
      return result;
    } catch (err) {
      lastErr = err;
      breaker?.recordFailure();
      const hasMoreAttempts = attempt < maxAttempts - 1;
      if (!hasMoreAttempts || !isRetriable(err)) throw err;
      // A server-stated `Retry-After` is a floor on the wait (never shorter than it asks, never
      // longer than the cap); the jittered backoff still applies when it says nothing.
      const stated = err instanceof HttpError ? err.retryAfterMs : undefined;
      await sleep(Math.max(backoffDelay(attempt, baseMs, capMs, random), stated ?? 0));
    }
  }
  // Unreachable: the loop either returns or throws. Present for exhaustiveness.
  throw lastErr;
}

/**
 * A transport-level error thrown by the read helpers to carry the HTTP status so the
 * retry predicate can classify it without re-parsing a message string. The body detail
 * is already truncated and secret-free by the caller.
 */
export class HttpError extends Error {
  constructor(
    readonly status: number,
    message: string,
    /** Server-stated minimum wait before retrying, already capped (see {@link parseRetryAfter}). */
    readonly retryAfterMs?: number,
  ) {
    super(message);
    this.name = 'HttpError';
  }
}

/** Thrown by {@link CircuitBreaker.assertClosed} while the circuit is open. */
export class CircuitOpenError extends Error {
  constructor(readonly retryAtMs: number) {
    super('circuit open: upstream temporarily unavailable');
    this.name = 'CircuitOpenError';
  }
}

/**
 * Default transient classifier: retry network throws (a `fetch` rejection — typically a
 * `TypeError`, never an `HttpError`) and retry HTTP 429 / 5xx. Never retry other 4xx —
 * those are deterministic (bad request, auth, not found) and a retry only wastes calls.
 */
export function defaultIsRetriable(err: unknown): boolean {
  if (err instanceof HttpError) {
    return err.status === 429 || (err.status >= 500 && err.status <= 599);
  }
  // A cancel (the user's, or the idle watchdog's) is final: retrying an aborted request only waits
  // out the backoff before failing again. An open circuit is a decision already made.
  if (err instanceof Error && (err.name === 'AbortError' || err instanceof CircuitOpenError))
    return false;
  // A non-HttpError throw from fetch is a network/transport failure → transient.
  return true;
}

export type CircuitState = 'closed' | 'open' | 'half-open';

/**
 * Minimal circuit breaker: opens after `threshold` failures within `windowMs` and rejects fast
 * until `cooldownMs` elapses, then allows a single half-open trial. A success closes it; a
 * failure re-opens it. Time is injectable for deterministic tests. Failures older than the
 * window no longer count, so a slow trickle of unrelated errors never trips it.
 */
export class CircuitBreaker {
  private failures: number[] = [];
  private openedAt = 0;
  private state: CircuitState = 'closed';

  constructor(
    private readonly threshold = 5,
    private readonly cooldownMs = 30_000,
    private readonly now: () => number = Date.now,
    private readonly windowMs = 60_000,
  ) {}

  get current(): CircuitState {
    this.refresh();
    return this.state;
  }

  /** Throw {@link CircuitOpenError} if the breaker is open and still cooling down. */
  assertClosed(): void {
    this.refresh();
    if (this.state === 'open') {
      throw new CircuitOpenError(this.openedAt + this.cooldownMs);
    }
  }

  recordSuccess(): void {
    this.failures = [];
    this.state = 'closed';
  }

  recordFailure(): void {
    const at = this.now();
    this.failures = [...this.failures.filter((t) => at - t < this.windowMs), at];
    // A half-open probe that fails re-opens at once; otherwise the window decides.
    if (this.state === 'half-open' || this.failures.length >= this.threshold) {
      this.state = 'open';
      this.openedAt = at;
      this.failures = [];
    }
  }

  private refresh(): void {
    if (this.state === 'open' && this.now() - this.openedAt >= this.cooldownMs) {
      this.state = 'half-open';
    }
  }
}
