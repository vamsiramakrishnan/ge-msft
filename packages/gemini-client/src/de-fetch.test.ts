import { describe, it, expect, vi } from 'vitest';
import { postJson } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/** A no-backoff retry policy so transient retries run synchronously in tests. */
const noSleep = { sleep: async () => {}, random: () => 0 };

function jsonResponse(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function tokenSource(overrides: Partial<TokenSource> = {}): TokenSource {
  return { getAccessToken: () => Promise.resolve('goog-token'), ...overrides };
}

describe('postJson', () => {
  it('sends the federated bearer token and JSON body, returns parsed JSON', async () => {
    const f = vi.fn(async () => jsonResponse({ ok: true, n: 7 }));
    const out = await postJson('https://de/api', { q: 'hi' }, tokenSource(), f as never);
    expect(out).toEqual({ ok: true, n: 7 });
    expect(f).toHaveBeenCalledTimes(1);
    const [url, init] = f.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://de/api');
    expect(init.method).toBe('POST');
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer goog-token');
    expect((init.headers as Record<string, string>)['Content-Type']).toBe('application/json');
    expect(JSON.parse(init.body as string)).toEqual({ q: 'hi' });
  });

  it('does not attach a signal when none is provided', async () => {
    const f = vi.fn(async () => jsonResponse({}));
    await postJson('https://de/api', {}, tokenSource(), f as never);
    const init = (f.mock.calls[0] as unknown as [string, RequestInit])[1];
    expect('signal' in init).toBe(false);
  });

  it('forwards the AbortSignal when provided', async () => {
    const f = vi.fn(async () => jsonResponse({}));
    const ctrl = new AbortController();
    await postJson('https://de/api', {}, tokenSource(), f as never, ctrl.signal);
    const init = (f.mock.calls[0] as unknown as [string, RequestInit])[1];
    expect(init.signal).toBe(ctrl.signal);
  });

  it('on 401 invalidates the token once and retries the send', async () => {
    const invalidate = vi.fn();
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(jsonResponse({ recovered: true }));
    const out = await postJson('https://de/api', {}, tokenSource({ invalidate }), f as never);
    expect(invalidate).toHaveBeenCalledOnce();
    expect(f).toHaveBeenCalledTimes(2);
    expect(out).toEqual({ recovered: true });
  });

  it('does NOT retry a 401 when the token source cannot invalidate', async () => {
    const f = vi.fn(async () => new Response('expired', { status: 401 }));
    // No invalidate() on the source → the 401 falls straight through to error formatting.
    await expect(postJson('https://de/api', {}, tokenSource(), f as never)).rejects.toThrow(
      /failed \(401\)/,
    );
    expect(f).toHaveBeenCalledTimes(1);
  });

  it('retries transient 5xx with backoff, then succeeds', async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('boom', { status: 503 }))
      .mockResolvedValueOnce(jsonResponse({ ok: 1 }));
    const out = await postJson('https://de/api', {}, tokenSource(), f as never, undefined, noSleep);
    expect(f).toHaveBeenCalledTimes(2);
    expect(out).toEqual({ ok: 1 });
  });

  it('falls through to error formatting when transient retries are exhausted', async () => {
    const f = vi.fn(async () => new Response('still down', { status: 500 }));
    await expect(
      postJson('https://de/api', {}, tokenSource(), f as never, undefined, {
        ...noSleep,
        maxAttempts: 2,
      }),
    ).rejects.toThrow(/https:\/\/de\/api failed \(500\): still down/);
    // 1 initial + 1 retry from maxAttempts:2.
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('propagates a genuine network/transport throw unchanged', async () => {
    const f = vi.fn(async () => {
      throw new TypeError('Failed to fetch');
    });
    await expect(
      postJson('https://de/api', {}, tokenSource(), f as never, undefined, {
        ...noSleep,
        maxAttempts: 1,
      }),
    ).rejects.toThrow(/Failed to fetch/);
  });

  it('formats a non-retriable 4xx error with the response body detail', async () => {
    const f = vi.fn(async () => new Response('bad filter', { status: 400 }));
    await expect(postJson('https://de/api', {}, tokenSource(), f as never)).rejects.toThrow(
      /https:\/\/de\/api failed \(400\): bad filter/,
    );
    // 400 is not transient → no retry.
    expect(f).toHaveBeenCalledTimes(1);
  });

  it('uses a generic message when the error response body cannot be read', async () => {
    // A Response whose body read throws → safeText returns '' → generic message branch.
    const broken = new Response('x', { status: 400 });
    Object.defineProperty(broken, 'text', {
      value: () => Promise.reject(new Error('stream torn')),
    });
    const f = vi.fn(async () => broken);
    await expect(postJson('https://de/api', {}, tokenSource(), f as never)).rejects.toThrow(
      /request failed \(400\)/,
    );
  });
});
