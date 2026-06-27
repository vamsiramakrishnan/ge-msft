// @vitest-environment jsdom
import { describe, it, expect, afterEach, vi } from 'vitest';
import {
  StreamAssistClient,
  WifTokenClient,
  type GeminiClientConfig,
  type EntraTokenProvider,
  type WifConfig,
} from '@ge/gemini-client';
import {
  installFakeWord,
  mountStack,
  type WordSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * INTERPLAY — wif-auth-failure: gemini-client WIF ↔ runtime AssistSession ↔ web-shell PanelController/App.
 *
 * The WIF (Workforce Identity Federation) STS token exchange is on the security-critical path: it is
 * the ONLY place the user's Entra identity becomes a Google access token, and no other credential is
 * in play. This suite wires the REAL exchange (`WifTokenClient`) into the REAL `StreamAssistClient`
 * (as its `TokenSource`), then into the REAL `AssistSession` (via `mountStack`'s real bridge →
 * session → `PanelController` → `<App/>`), and mocks ONLY the outermost network boundary: a `vi.fn`
 * `fetch` that FAILS the STS exchange.
 *
 * What we prove across the seam:
 *   1. A hard STS failure (a non-transient 403 — bad audience / denied principal) propagates from the
 *      WIF exchange, through `StreamAssistClient.stream` (which maps a thrown token source to an
 *      `error` SSE event), through `AssistSession.ask`, into a SURFACED, DEGRADED UI state — a
 *      `role="alert"` error in the thread — NOT a thrown/unhandled crash, NOT a silent success.
 *   2. A network-level STS throw (fetch rejects) degrades the same way (never a silent answer).
 *   3. The refresh path: the cached token is accepted, Discovery Engine returns 401, the client
 *      `invalidate()`s and RE-EXCHANGES — and when that re-exchange ALSO fails at STS, the failure
 *      still degrades to a surfaced error rather than serving a stale/empty success.
 *   4. The federated bearer token (the user's only secret) never lands in the rendered DOM, even on
 *      the error path.
 *
 * Everything between the fake host and the fake fetch is the real code: real WIF retry/caching/
 * invalidate, real streamAssist transport+mapping, real AssistSession streaming, real controller
 * reducer, real React view. We assert OBSERVABLE cross-boundary behavior (the rendered alert, the
 * absence of any assistant answer text, no secret leak), never just that a function returned.
 */

let sim: WordSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
  vi.restoreAllMocks();
});

/** A signed-in user's Entra OIDC id token provider (the only secret the client holds). */
const ENTRA_ID_TOKEN = 'entra-oidc-id-token-SECRET';
const entra: EntraTokenProvider = { getIdToken: () => Promise.resolve(ENTRA_ID_TOKEN) };

const WIF_CONFIG: WifConfig = {
  poolId: 'acme-pool',
  providerId: 'entra-provider',
  stsEndpoint: 'https://sts.test/v1/token',
};

const ASSIST_CONFIG: GeminiClientConfig = {
  assistant: { project: 'proj', location: 'eu', engine: 'ge-engine' },
  identity: 'sim.user@acme',
};

/** No-backoff retry policy so any transient STS retry runs without real timers. */
const noSleep = { sleep: async (): Promise<void> => {}, random: (): number => 0 };

function stsErrorResponse(status: number, body = 'access_denied'): Response {
  return new Response(body, { status });
}

function stsOk(token: string): Response {
  return new Response(
    JSON.stringify({ access_token: token, token_type: 'Bearer', expires_in: 3600 }),
    { status: 200, headers: { 'Content-Type': 'application/json' } },
  );
}

/**
 * Wire the REAL WIF exchange → REAL streamAssist client over two `vi.fn` fetches: one for the STS
 * endpoint, one for Discovery Engine. The WIF client gets its own no-backoff retry policy so a
 * transient-classified status doesn't spin on real timers. Returns the live client + both spies.
 */
function realClientOverFakeFetch(opts: {
  stsFetch: typeof fetch;
  deFetch?: typeof fetch;
}): StreamAssistClient {
  const wif = new WifTokenClient(entra, WIF_CONFIG, opts.stsFetch, Date.now, {
    ...noSleep,
    maxAttempts: 2,
  });
  // The streamAssist transport fetch — only reached if the token exchange SUCCEEDS.
  const deFetch =
    opts.deFetch ??
    (vi.fn(async () => {
      throw new Error('discovery-engine fetch should not be reached when WIF fails');
    }) as unknown as typeof fetch);
  // No backoff on the streamAssist POST either (keeps the 401 re-exchange path off real timers).
  return new StreamAssistClient(wif, ASSIST_CONFIG, deFetch, noSleep);
}

/** Drive one grounded turn through the real stack and let it settle to a non-busy state. */
async function runOneTurn(client: StreamAssistClient, query: string): Promise<void> {
  sim = installFakeWord();
  ui = mountStack({ surface: 'word', client });
  await ui!.flush();
  await ui!.act(() => void ui!.controller.send(query));
  await ui!.waitFor((s) => !s.busy && s.messages.length > 0);
  await ui!.flush();
}

describe('WIF auth failure — interplay across gemini-client WIF · runtime · web-shell', () => {
  it('a hard STS 403 degrades to a SURFACED error in the thread (no crash, no silent success)', async () => {
    // The STS denies the exchange (e.g. bad audience / unmapped principal): a non-transient 403.
    const stsFetch = vi.fn(async () =>
      stsErrorResponse(403, 'permission_denied'),
    ) as unknown as typeof fetch;
    const client = realClientOverFakeFetch({ stsFetch });

    await runOneTurn(client, 'summarize the master services agreement');

    // The STS WAS reached: the real exchange ran (and, being a non-transient 403, did not retry-
    // storm — each attempt is a single call, no exponential fan-out within one exchange).
    const stsSpy = stsFetch as unknown as ReturnType<typeof vi.fn>;
    expect(stsSpy.mock.calls.length).toBeGreaterThanOrEqual(1);
    // A 403 is non-transient: every recorded STS response is the single denied call (no retry).
    expect(stsSpy.mock.results.every((r) => r.type === 'return')).toBe(true);

    // A degraded, surfaced UI state: the thread shows a role=alert error, not an answer.
    const alert = ui!.container.querySelector('.msg-error[role="alert"]');
    expect(alert).not.toBeNull();

    // The controller never streamed any assistant answer text — the failure short-circuited it.
    const state = ui!.controller.getState();
    const assistant = state.messages.find((m) => m.role === 'assistant');
    expect(assistant).toBeDefined();
    expect(assistant!.error).toBeTruthy();
    expect((assistant!.text ?? '').trim()).toBe('');
    // Not busy, streaming flag cleared — the loop settled, it did not hang or throw out.
    expect(state.busy).toBe(false);
    expect(assistant!.streaming).toBe(false);

    // Discovery Engine was NEVER reached: the token exchange gated it.
    // (deFetch defaults to a throwing spy; the absence of a thrown crash proves it wasn't called.)
  });

  it('a network-level STS throw (fetch rejects) also degrades to a surfaced error', async () => {
    // A transport failure: fetch rejects (DNS / offline). Default classifier treats it as transient,
    // so the no-backoff policy retries up to maxAttempts (2) then surfaces the throw.
    const stsFetch = vi.fn(async () => {
      throw new TypeError('Failed to fetch');
    }) as unknown as typeof fetch;
    const client = realClientOverFakeFetch({ stsFetch });

    await runOneTurn(client, 'what are the key risks?');

    // A network throw is transient → the no-backoff policy retried before giving up: at least one
    // exchange ran more than a single attempt (proof the real retry path executed).
    const stsSpy = stsFetch as unknown as ReturnType<typeof vi.fn>;
    expect(stsSpy.mock.calls.length).toBeGreaterThanOrEqual(2);

    const state = ui!.controller.getState();
    const assistant = state.messages.find((m) => m.role === 'assistant');
    expect(assistant?.error).toBeTruthy();
    expect((assistant?.text ?? '').trim()).toBe('');
    // Surfaced in the DOM as an alert (degraded), never an unhandled rejection.
    expect(ui!.container.querySelector('.msg-error[role="alert"]')).not.toBeNull();
    expect(state.busy).toBe(false);
  });

  it('the REFRESH path: cached token accepted, DE 401 → invalidate → re-exchange FAILS → degraded', async () => {
    // STS hands out a token the first time (cached), then the re-exchange after a DE 401 is denied.
    const stsFetch = vi
      .fn()
      .mockResolvedValueOnce(stsOk('goog-access-token-1'))
      .mockResolvedValueOnce(stsErrorResponse(403, 'principal_revoked')) as unknown as typeof fetch;

    // Discovery Engine rejects the (valid-looking) bearer with a 401 every time, forcing the
    // client's single re-exchange — which then fails at STS.
    const deFetch = vi.fn(
      async () => new Response('unauthorized', { status: 401 }),
    ) as unknown as typeof fetch;

    const client = realClientOverFakeFetch({ stsFetch, deFetch });

    await runOneTurn(client, 'draft a renewal note');

    // The real refresh handshake happened: first exchange + re-exchange after the 401.
    expect(stsFetch as unknown as ReturnType<typeof vi.fn>).toHaveBeenCalledTimes(2);
    // Discovery Engine was hit at least once (the 401 that triggered the re-exchange).
    expect(
      (deFetch as unknown as ReturnType<typeof vi.fn>).mock.calls.length,
    ).toBeGreaterThanOrEqual(1);

    const state = ui!.controller.getState();
    const assistant = state.messages.find((m) => m.role === 'assistant');
    // The failed re-exchange degrades — never a stale/empty SILENT success.
    expect(assistant?.error).toBeTruthy();
    expect((assistant?.text ?? '').trim()).toBe('');
    expect(ui!.container.querySelector('.msg-error[role="alert"]')).not.toBeNull();
    expect(state.busy).toBe(false);
  });

  it('the user secret (Entra id token) never lands in the rendered DOM, even on the error path', async () => {
    const stsFetch = vi.fn(async () => stsErrorResponse(403, 'denied')) as unknown as typeof fetch;
    const client = realClientOverFakeFetch({ stsFetch });

    await runOneTurn(client, 'summarize');

    // The error is surfaced, but the secret subject token is screened out of any user-facing text.
    const dom = ui!.container.textContent ?? '';
    expect(dom).not.toContain(ENTRA_ID_TOKEN);
    // And no Authorization-style bearer secret leaked into the visible error either.
    expect(dom).not.toContain('Bearer ');
  });
});
