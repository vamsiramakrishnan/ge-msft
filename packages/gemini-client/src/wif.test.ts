import { describe, it, expect, vi } from 'vitest';
import { WifTokenClient } from './wif.js';

const entra = { getIdToken: () => Promise.resolve('entra-id-token') };

/** No-backoff retry policy so transient STS retries run without real timers. */
const noSleep = { sleep: async () => {}, random: () => 0 };

function jsonResponse(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function stsOk(token: string, expiresIn = 3600): Response {
  return jsonResponse({ access_token: token, token_type: 'Bearer', expires_in: expiresIn });
}

describe('WifTokenClient — request shaping', () => {
  it('uses a receiver-safe default fetch implementation in browser frames', async () => {
    const receiverSensitiveFetch = vi.fn(function (
      this: unknown,
      _input: RequestInfo | URL,
      _init?: RequestInit,
    ) {
      if (this !== globalThis) throw new TypeError('Illegal invocation');
      return Promise.resolve(stsOk('goog-bound-fetch'));
    });
    vi.stubGlobal('fetch', receiverSensitiveFetch);
    try {
      const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' });
      expect(await wif.getAccessToken()).toBe('goog-bound-fetch');
      expect(receiverSensitiveFetch).toHaveBeenCalledWith(
        'https://sts.googleapis.com/v1/token',
        expect.objectContaining({ method: 'POST' }),
      );
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('passes scope override and userProject options in the exchange body', async () => {
    const f = vi.fn(async () => stsOk('goog-1'));
    const wif = new WifTokenClient(
      entra,
      {
        poolId: 'p',
        providerId: 'pr',
        scope: 'https://www.googleapis.com/auth/discoveryengine',
        userProject: 'billing-proj',
      },
      f as never,
    );
    await wif.getAccessToken();
    const body = JSON.parse(
      (f.mock.calls[0] as unknown as [string, RequestInit])[1].body as string,
    );
    expect(body.scope).toBe('https://www.googleapis.com/auth/discoveryengine');
    expect(JSON.parse(body.options)).toEqual({ userProject: 'billing-proj' });
    expect(body.grantType).toBe('urn:ietf:params:oauth:grant-type:token-exchange');
    expect(body.requestedTokenType).toBe('urn:ietf:params:oauth:token-type:access_token');
    expect(body.subjectTokenType).toBe('urn:ietf:params:oauth:token-type:id_token');
  });

  it('defaults to the cloud-platform scope and omits options when no userProject', async () => {
    const f = vi.fn(async () => stsOk('goog-1'));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    await wif.getAccessToken();
    const body = JSON.parse(
      (f.mock.calls[0] as unknown as [string, RequestInit])[1].body as string,
    );
    expect(body.scope).toBe('https://www.googleapis.com/auth/cloud-platform');
    expect('options' in body).toBe(false);
  });

  it('targets the override STS endpoint when configured', async () => {
    const f = vi.fn(async () => stsOk('goog-1'));
    const wif = new WifTokenClient(
      entra,
      { poolId: 'p', providerId: 'pr', stsEndpoint: 'https://sts.test/exchange' },
      f as never,
    );
    await wif.getAccessToken();
    expect((f.mock.calls[0] as unknown as [string])[0]).toBe('https://sts.test/exchange');
  });
});

describe('WifTokenClient — caching and TTL', () => {
  it('defaults TTL to 3600s when STS omits expires_in', async () => {
    let now = 0;
    const f = vi.fn(async () => jsonResponse({ access_token: 'goog-1', token_type: 'Bearer' }));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never, () => now);
    expect(await wif.getAccessToken()).toBe('goog-1');
    // Just under the 3600s TTL (minus 60s skew) → still cached.
    now = (3600 - 60 - 1) * 1000;
    expect(await wif.getAccessToken()).toBe('goog-1');
    expect(f).toHaveBeenCalledTimes(1);
  });

  it('re-exchanges within the skew window before actual expiry', async () => {
    let now = 0;
    const f = vi
      .fn()
      .mockResolvedValueOnce(stsOk('goog-A', 100))
      .mockResolvedValueOnce(stsOk('goog-B', 100));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never, () => now);
    expect(await wif.getAccessToken()).toBe('goog-A');
    // 50s in: TTL 100s minus 60s skew means it is already due for refresh.
    now = 50_000;
    expect(await wif.getAccessToken()).toBe('goog-B');
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('re-exchanges after invalidate() (e.g. a 401 from Discovery Engine)', async () => {
    const f = vi.fn().mockResolvedValueOnce(stsOk('goog-A')).mockResolvedValueOnce(stsOk('goog-B'));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    expect(await wif.getAccessToken()).toBe('goog-A');
    wif.invalidate();
    expect(await wif.getAccessToken()).toBe('goog-B');
    expect(f).toHaveBeenCalledTimes(2);
  });
});

describe('WifTokenClient — transient failure handling', () => {
  it('retries a transient 503 STS failure with backoff, then succeeds', async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('unavailable', { status: 503 }))
      .mockResolvedValueOnce(stsOk('goog-after-retry'));
    const wif = new WifTokenClient(
      entra,
      { poolId: 'p', providerId: 'pr' },
      f as never,
      Date.now,
      noSleep,
    );
    expect(await wif.getAccessToken()).toBe('goog-after-retry');
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('retries a 429 and surfaces the secret-free status in the retriable error message', async () => {
    // All attempts 429 → exhausted → the last thrown HttpError carries the formatted detail.
    const f = vi.fn(async () => new Response('Too Many Requests', { status: 429 }));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never, Date.now, {
      ...noSleep,
      maxAttempts: 2,
    });
    await expect(wif.getAccessToken()).rejects.toThrow(/WIF token exchange failed \(429\)/);
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('does not retry a non-transient 400 (bad audience) — throws on the first attempt', async () => {
    const f = vi.fn(async () => new Response('invalid_grant', { status: 400 }));
    const wif = new WifTokenClient(
      entra,
      { poolId: 'p', providerId: 'pr' },
      f as never,
      Date.now,
      noSleep,
    );
    await expect(wif.getAccessToken()).rejects.toThrow(
      /WIF token exchange failed \(400\): invalid_grant/,
    );
    expect(f).toHaveBeenCalledTimes(1);
  });

  it('uses a no-body marker when the failed STS response body cannot be read', async () => {
    const broken = new Response('x', { status: 400 });
    Object.defineProperty(broken, 'text', {
      value: () => Promise.reject(new Error('torn')),
    });
    const f = vi.fn(async () => broken);
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    await expect(wif.getAccessToken()).rejects.toThrow(/<no body>/);
  });

  it('does not cache a token after a failed exchange — next call re-exchanges', async () => {
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('bad', { status: 400 }))
      .mockResolvedValueOnce(stsOk('goog-recovered'));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    await expect(wif.getAccessToken()).rejects.toThrow(/400/);
    // The inflight promise must have cleared so a second call can succeed.
    expect(await wif.getAccessToken()).toBe('goog-recovered');
  });

  it('rejects an STS response missing access_token (schema parse failure)', async () => {
    const f = vi.fn(async () => jsonResponse({ token_type: 'Bearer', expires_in: 3600 }));
    const wif = new WifTokenClient(entra, { poolId: 'p', providerId: 'pr' }, f as never);
    await expect(wif.getAccessToken()).rejects.toThrow();
  });
});

describe('WifTokenClient — stale id token recovery', () => {
  const STALE = JSON.stringify({
    error: 'invalid_grant',
    error_description: 'ID Token issued at 1790788565 is stale to sign-in.',
  });

  it('re-acquires a freshly issued id token once when the STS calls the cached one stale', async () => {
    // Live 2026-10-01: MSAL kept serving a cached id token the STS rejected, so every request
    // failed until sign-in, and reloading the task pane did not help.
    const getIdToken = vi.fn(async (opts?: { forceRefresh?: boolean }) =>
      opts?.forceRefresh ? 'fresh-id-token' : 'cached-id-token',
    );
    const bodies: string[] = [];
    const f = vi.fn(async (_url: string, init?: { body?: string }) => {
      bodies.push(init?.body ?? '');
      return bodies.length === 1 ? new Response(STALE, { status: 400 }) : stsOk('goog-fresh');
    });
    const wif = new WifTokenClient(
      { getIdToken },
      { poolId: 'p', providerId: 'pr' },
      f as never,
      Date.now,
      noSleep,
    );
    expect(await wif.getAccessToken()).toBe('goog-fresh');
    expect(getIdToken).toHaveBeenNthCalledWith(2, { forceRefresh: true });
    expect(JSON.parse(bodies[1]!).subjectToken).toBe('fresh-id-token');
  });

  it('retries only once, and never for other 400s', async () => {
    const getIdToken = vi.fn(async () => 'id-token');
    const stale = vi.fn(async () => new Response(STALE, { status: 400 }));
    const a = new WifTokenClient(
      { getIdToken },
      { poolId: 'p', providerId: 'pr' },
      stale as never,
      Date.now,
      noSleep,
    );
    await expect(a.getAccessToken()).rejects.toThrow(/stale to sign-in/);
    expect(stale).toHaveBeenCalledTimes(2);

    const badAudience = vi.fn(
      async () =>
        new Response('{"error":"invalid_grant","error_description":"bad audience"}', {
          status: 400,
        }),
    );
    const b = new WifTokenClient(
      { getIdToken },
      { poolId: 'p', providerId: 'pr' },
      badAudience as never,
      Date.now,
      noSleep,
    );
    await expect(b.getAccessToken()).rejects.toThrow(/bad audience/);
    expect(badAudience).toHaveBeenCalledTimes(1);
  });

  it('hands a token fetched across invalidate() to the caller without caching it', async () => {
    const holder: { wif?: WifTokenClient } = {};
    const getIdToken = vi.fn(async (opts?: { forceRefresh?: boolean }) => {
      if (opts?.forceRefresh) holder.wif?.invalidate(); // an invalidate lands during the forced refresh
      return 'id-token';
    });
    let call = 0;
    const f = vi.fn(async () => {
      call += 1;
      if (call === 1) return new Response(STALE, { status: 400 });
      return stsOk(`goog-${call}`);
    });
    const wif = new WifTokenClient(
      { getIdToken },
      { poolId: 'p', providerId: 'pr' },
      f as never,
      Date.now,
      noSleep,
    );
    holder.wif = wif;
    expect(await wif.getAccessToken()).toBe('goog-2');
    expect(await wif.getAccessToken()).toBe('goog-3'); // not served from cache
  });

  it('marks a failed exchange as a non-retriable HTTP error for outer retry policies', async () => {
    const { defaultIsRetriable } = await import('./retry.js');
    const f = vi.fn(
      async () =>
        new Response('{"error":"invalid_grant","error_description":"bad audience"}', {
          status: 400,
        }),
    );
    const wif = new WifTokenClient(
      entra,
      { poolId: 'p', providerId: 'pr' },
      f as never,
      Date.now,
      noSleep,
    );
    const err = await wif.getAccessToken().catch((e: unknown) => e);
    expect(defaultIsRetriable(err)).toBe(false);
  });
});
