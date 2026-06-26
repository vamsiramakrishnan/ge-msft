import type { TokenSource } from './stream-assist.js';
import { withRetry, defaultIsRetriable, HttpError, type RetryOptions } from './retry.js';

export type FetchLike = typeof fetch;

export const defaultFetch: FetchLike = (input, init) => globalThis.fetch(input, init);

/**
 * POST a JSON body as the signed-in user and return the parsed JSON.
 * Mirrors StreamAssistClient.post: federated bearer token, and on a 401 (token
 * expired mid-cache) invalidate once and re-exchange before retrying.
 * No Google credential ever reaches the client — only the federated token from the
 * TokenSource is used; the body is sent as data.
 *
 * These calls are idempotent reads, so we wrap the send in exponential backoff (full
 * jitter) for transient failures — network throws and HTTP 429/5xx. The existing
 * 401→invalidate→retry-once stays a *distinct* concern: a 401 is handled by re-exchanging
 * the token once and never rides the backoff path.
 */
export async function postJson(
  url: string,
  body: unknown,
  tokens: TokenSource,
  fetchImpl: FetchLike,
  signal?: AbortSignal,
  retryOpts: RetryOptions = {},
): Promise<unknown> {
  const payload = JSON.stringify(body);
  const send = async (): Promise<Response> => {
    const token = await tokens.getAccessToken();
    return fetchImpl(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: payload,
      ...(signal ? { signal } : {}),
    });
  };

  // Transient-only backoff around the send; a 401 is excluded so it can take the
  // dedicated re-exchange path below rather than being retried with jitter. We keep the
  // last Response so an exhausted-but-still-transient outcome flows into the unified
  // error formatting below instead of throwing a bare HttpError.
  let lastRes: Response | undefined;
  const sendWithBackoff = async (): Promise<Response> => {
    const res = await send();
    lastRes = res;
    if (res.status !== 401 && defaultIsRetriable(new HttpError(res.status, ''))) {
      throw new HttpError(res.status, `request failed (${res.status})`);
    }
    return res;
  };

  let res: Response;
  try {
    res = await withRetry(sendWithBackoff, retryOpts);
  } catch (err) {
    if (err instanceof HttpError && lastRes) {
      res = lastRes; // exhausted transient → fall through to the shared error formatting.
    } else {
      throw err; // genuine network/transport throw → propagate unchanged.
    }
  }

  if (res.status === 401 && tokens.invalidate) {
    tokens.invalidate();
    res = await send();
  }
  if (!res.ok) {
    const detail = await safeText(res);
    throw new Error(
      detail ? `${url} failed (${res.status}): ${detail}` : `request failed (${res.status})`,
    );
  }
  return res.json();
}

export async function postJsonWithHeaders(
  url: string,
  body: unknown,
  tokens: TokenSource,
  fetchImpl: FetchLike,
  headers: Record<string, string>,
  signal?: AbortSignal,
  retryOpts: RetryOptions = {},
): Promise<unknown> {
  const payload = JSON.stringify(body);
  const send = async (): Promise<Response> => {
    const token = await tokens.getAccessToken();
    return fetchImpl(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...headers,
      },
      body: payload,
      ...(signal ? { signal } : {}),
    });
  };

  let lastRes: Response | undefined;
  const sendWithBackoff = async (): Promise<Response> => {
    const res = await send();
    lastRes = res;
    if (res.status !== 401 && defaultIsRetriable(new HttpError(res.status, ''))) {
      throw new HttpError(res.status, `request failed (${res.status})`);
    }
    return res;
  };

  let res: Response;
  try {
    res = await withRetry(sendWithBackoff, retryOpts);
  } catch (err) {
    if (err instanceof HttpError && lastRes) res = lastRes;
    else throw err;
  }

  if (res.status === 401 && tokens.invalidate) {
    tokens.invalidate();
    res = await send();
  }
  if (!res.ok) {
    const detail = await safeText(res);
    throw new Error(
      detail ? `${url} failed (${res.status}): ${detail}` : `request failed (${res.status})`,
    );
  }
  return res.json();
}

/**
 * GET JSON as the signed-in user. Used for browser-safe Discovery Engine catalog discovery
 * (assistant agents/skills and data stores). Same token posture as postJson: memory-only
 * federated bearer token, 401 invalidates once, bounded retries only for idempotent reads.
 */
export async function getJson(
  url: string,
  tokens: TokenSource,
  fetchImpl: FetchLike,
  signal?: AbortSignal,
  retryOpts: RetryOptions = {},
): Promise<unknown> {
  const send = async (): Promise<Response> => {
    const token = await tokens.getAccessToken();
    return fetchImpl(url, {
      method: 'GET',
      headers: { Authorization: `Bearer ${token}` },
      ...(signal ? { signal } : {}),
    });
  };

  let lastRes: Response | undefined;
  const sendWithBackoff = async (): Promise<Response> => {
    const res = await send();
    lastRes = res;
    if (res.status !== 401 && defaultIsRetriable(new HttpError(res.status, ''))) {
      throw new HttpError(res.status, `request failed (${res.status})`);
    }
    return res;
  };

  let res: Response;
  try {
    res = await withRetry(sendWithBackoff, retryOpts);
  } catch (err) {
    if (err instanceof HttpError && lastRes) res = lastRes;
    else throw err;
  }

  if (res.status === 401 && tokens.invalidate) {
    tokens.invalidate();
    res = await send();
  }
  if (!res.ok) {
    const detail = await safeText(res);
    throw new Error(
      detail ? `${url} failed (${res.status}): ${detail}` : `request failed (${res.status})`,
    );
  }
  return res.json();
}

async function safeText(res: Response): Promise<string> {
  try {
    return (await res.text()).slice(0, 500);
  } catch {
    return '';
  }
}
