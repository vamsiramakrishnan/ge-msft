import { describe, it, expect, vi } from 'vitest';
import { withRetry, backoffDelay, defaultIsRetriable, HttpError, CircuitBreaker } from './retry.js';

/** A deterministic sleep that records every delay it was asked to wait. */
function recordingSleep() {
  const delays: number[] = [];
  const sleep = vi.fn(async (ms: number) => {
    delays.push(ms);
  });
  return { delays, sleep };
}

describe('defaultIsRetriable', () => {
  it('retries 429 and 5xx but not other 4xx', () => {
    expect(defaultIsRetriable(new HttpError(429, ''))).toBe(true);
    expect(defaultIsRetriable(new HttpError(500, ''))).toBe(true);
    expect(defaultIsRetriable(new HttpError(503, ''))).toBe(true);
    expect(defaultIsRetriable(new HttpError(400, ''))).toBe(false);
    expect(defaultIsRetriable(new HttpError(401, ''))).toBe(false);
    expect(defaultIsRetriable(new HttpError(404, ''))).toBe(false);
  });
  it('retries a non-HttpError throw (network/transport failure)', () => {
    expect(defaultIsRetriable(new TypeError('Failed to fetch'))).toBe(true);
  });
});

describe('backoffDelay (full jitter)', () => {
  it('draws from [0, min(cap, base*2**attempt))', () => {
    // random()=0.5, base=100: attempt0→[0,100)→50, attempt1→[0,200)→100, attempt2→[0,400)→200
    expect(backoffDelay(0, 100, 5000, () => 0.5)).toBe(50);
    expect(backoffDelay(1, 100, 5000, () => 0.5)).toBe(100);
    expect(backoffDelay(2, 100, 5000, () => 0.5)).toBe(200);
  });
  it('caps the window', () => {
    expect(backoffDelay(10, 100, 1000, () => 0.999)).toBeLessThan(1000);
  });
});

describe('withRetry', () => {
  it('retries a 429 then succeeds, backing off with the injected sleep/random', async () => {
    const { delays, sleep } = recordingSleep();
    const fn = vi
      .fn()
      .mockRejectedValueOnce(new HttpError(429, 'rate limited'))
      .mockResolvedValueOnce('ok');
    const out = await withRetry(fn, { sleep, random: () => 0.5, baseMs: 100 });
    expect(out).toBe('ok');
    expect(fn).toHaveBeenCalledTimes(2);
    expect(delays).toEqual([50]); // one backoff of base*2**0 * 0.5
  });

  it('retries a 5xx and a network throw before succeeding', async () => {
    const { delays, sleep } = recordingSleep();
    const fn = vi
      .fn()
      .mockRejectedValueOnce(new HttpError(503, ''))
      .mockRejectedValueOnce(new TypeError('network down'))
      .mockResolvedValueOnce('done');
    const out = await withRetry(fn, { sleep, random: () => 1, baseMs: 10, maxAttempts: 5 });
    expect(out).toBe('done');
    expect(fn).toHaveBeenCalledTimes(3);
    expect(delays).toHaveLength(2);
  });

  it('does NOT retry a 4xx (throws immediately, no sleep)', async () => {
    const { delays, sleep } = recordingSleep();
    const fn = vi.fn().mockRejectedValue(new HttpError(400, 'bad'));
    await expect(withRetry(fn, { sleep })).rejects.toBeInstanceOf(HttpError);
    expect(fn).toHaveBeenCalledTimes(1);
    expect(delays).toEqual([]);
  });

  it('caps total attempts and rethrows the last error', async () => {
    const { delays, sleep } = recordingSleep();
    const err = new HttpError(500, 'boom');
    const fn = vi.fn().mockRejectedValue(err);
    await expect(withRetry(fn, { sleep, maxAttempts: 3, random: () => 0 })).rejects.toBe(err);
    expect(fn).toHaveBeenCalledTimes(3); // 1 initial + 2 retries
    expect(delays).toHaveLength(2); // a sleep before each retry, none after the last failure
  });
});

describe('CircuitBreaker', () => {
  it('opens after threshold consecutive failures and half-opens after cooldown', () => {
    let now = 0;
    const cb = new CircuitBreaker(2, 1000, () => now);
    expect(cb.current).toBe('closed');
    cb.recordFailure();
    cb.recordFailure();
    expect(cb.current).toBe('open');
    expect(() => cb.assertClosed()).toThrow(/circuit open/);
    now += 1000;
    expect(cb.current).toBe('half-open');
    cb.recordSuccess();
    expect(cb.current).toBe('closed');
  });
});
