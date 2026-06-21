import type { TokenSource } from './stream-assist.js';

export type FetchLike = typeof fetch;

/**
 * POST a JSON body as the signed-in user and return the parsed JSON.
 * Mirrors StreamAssistClient.post: federated bearer token, and on a 401 (token
 * expired mid-cache) invalidate once and re-exchange before retrying.
 * Google credentials never leave the gateway/token source; the body is sent as data.
 */
export async function postJson(
  url: string,
  body: unknown,
  tokens: TokenSource,
  fetchImpl: FetchLike,
  signal?: AbortSignal,
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

  let res = await send();
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
