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
      await sleep(backoffDelay(attempt, baseMs, capMs, random));
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
  ) {
    super(message);
    this.name = 'HttpError';
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
  // A non-HttpError throw from fetch is a network/transport failure → transient.
  return true;
}

export type CircuitState = 'closed' | 'open' | 'half-open';

/**
 * Minimal circuit breaker: opens after `threshold` consecutive failures and rejects
 * fast until `cooldownMs` elapses, then allows a single half-open trial. A success
 * closes it; a failure re-opens it. Time is injectable for deterministic tests.
 */
export class CircuitBreaker {
  private failures = 0;
  private openedAt = 0;
  private state: CircuitState = 'closed';

  constructor(
    private readonly threshold = 5,
    private readonly cooldownMs = 30_000,
    private readonly now: () => number = Date.now,
  ) {}

  get current(): CircuitState {
    this.refresh();
    return this.state;
  }

  /** Throw if the breaker is open and still cooling down. */
  assertClosed(): void {
    this.refresh();
    if (this.state === 'open') {
      throw new Error('circuit open: upstream temporarily unavailable');
    }
  }

  recordSuccess(): void {
    this.failures = 0;
    this.state = 'closed';
  }

  recordFailure(): void {
    this.failures += 1;
    if (this.failures >= this.threshold) {
      this.state = 'open';
      this.openedAt = this.now();
    }
  }

  private refresh(): void {
    if (this.state === 'open' && this.now() - this.openedAt >= this.cooldownMs) {
      this.state = 'half-open';
    }
  }
}
