import { z } from 'zod';

/**
 * Workforce Identity Federation token exchange, run in the browser/webview.
 * The user's Entra OIDC token is exchanged at Google's STS for a short-lived
 * Google access token (RFC 8693). No Google service-account key is ever held by
 * the client — the only secret in play is the user's own short-lived Entra token.
 * See ADR-0001 and docs/api/discoveryengine/README.md.
 */

/** Supplies the signed-in user's Entra **OIDC id token** (e.g. from MSAL NAA). */
export interface EntraTokenProvider {
  getIdToken(): Promise<string>;
}

export interface WifConfig {
  poolId: string; // Workforce Identity Pool id
  providerId: string; // the Entra provider id within the pool
  /** Defaults to cloud-platform; narrow if your engine accepts a tighter scope. */
  scope?: string;
  /** Optional billing/quota project passed via STS `options.userProject`. */
  userProject?: string;
  /** Override for tests. */
  stsEndpoint?: string;
}

const STS_ENDPOINT = 'https://sts.googleapis.com/v1/token';
const TOKEN_EXCHANGE_GRANT = 'urn:ietf:params:oauth:grant-type:token-exchange';
const ID_TOKEN_TYPE = 'urn:ietf:params:oauth:token-type:id_token';
const ACCESS_TOKEN_TYPE = 'urn:ietf:params:oauth:token-type:access_token';
const CLOUD_PLATFORM_SCOPE = 'https://www.googleapis.com/auth/cloud-platform';

const StsResponseSchema = z.object({
  access_token: z.string(),
  token_type: z.string(),
  expires_in: z.number().optional(), // seconds
});

/** Refresh this many seconds before actual expiry to avoid mid-flight 401s. */
const EXPIRY_SKEW_SECONDS = 60;

type FetchLike = typeof fetch;

interface CachedToken {
  accessToken: string;
  expiresAtMs: number;
}

export class WifTokenClient {
  private cached: CachedToken | null = null;
  private inflight: Promise<string> | null = null;
  /** Bumped by invalidate(); an exchange only commits its result if this is unchanged. */
  private epoch = 0;

  constructor(
    private readonly entra: EntraTokenProvider,
    private readonly config: WifConfig,
    private readonly fetchImpl: FetchLike = fetch,
    private readonly now: () => number = Date.now,
  ) {}

  private audience(): string {
    return (
      `//iam.googleapis.com/locations/global/workforcePools/` +
      `${this.config.poolId}/providers/${this.config.providerId}`
    );
  }

  /** Returns a valid Google access token, reusing the cached one until near expiry. */
  async getAccessToken(): Promise<string> {
    if (this.cached && this.cached.expiresAtMs - EXPIRY_SKEW_SECONDS * 1000 > this.now()) {
      return this.cached.accessToken;
    }
    // Collapse concurrent refreshes into one exchange.
    if (this.inflight) return this.inflight;
    this.inflight = this.exchange().finally(() => {
      this.inflight = null;
    });
    return this.inflight;
  }

  /** Force the next call to re-exchange (e.g. after a 401 from Discovery Engine). */
  invalidate(): void {
    this.cached = null;
    // Bump the epoch so any inflight exchange won't reinstate a pre-invalidate token.
    this.epoch += 1;
  }

  private async exchange(): Promise<string> {
    const startEpoch = this.epoch;
    const idToken = await this.entra.getIdToken();
    const body: Record<string, string> = {
      grantType: TOKEN_EXCHANGE_GRANT,
      audience: this.audience(),
      scope: this.config.scope ?? CLOUD_PLATFORM_SCOPE,
      requestedTokenType: ACCESS_TOKEN_TYPE,
      subjectToken: idToken,
      subjectTokenType: ID_TOKEN_TYPE,
    };
    if (this.config.userProject) {
      body.options = JSON.stringify({ userProject: this.config.userProject });
    }

    const res = await this.fetchImpl(this.config.stsEndpoint ?? STS_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const detail = await safeText(res);
      throw new Error(`WIF token exchange failed (${res.status}): ${detail}`);
    }
    const parsed = StsResponseSchema.parse(await res.json());
    // If invalidate() ran while this exchange was in flight, hand the freshly-fetched
    // token to *this* caller but do not cache it — the next call must re-exchange.
    if (this.epoch === startEpoch) {
      const ttlMs = (parsed.expires_in ?? 3600) * 1000;
      this.cached = { accessToken: parsed.access_token, expiresAtMs: this.now() + ttlMs };
    }
    return parsed.access_token;
  }
}

async function safeText(res: Response): Promise<string> {
  try {
    return (await res.text()).slice(0, 500);
  } catch {
    return '<no body>';
  }
}
