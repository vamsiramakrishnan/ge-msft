import type { EntraTokenProvider } from '@ge/gemini-client';
import type { AuthClient, UserIdentity } from '@ge/runtime';

/**
 * The signed-in user's identity envelope, backed by MSAL **Nested App Authentication (NAA)**.
 * It supplies three things and never touches a Google credential:
 *   • `getIdToken()`  — the Entra OIDC id token used as the WIF subject (→ Google access token).
 *   • `getGraphToken()` — a delegated Microsoft Graph access token (Plane B / estate reads).
 *   • `getIdentity()` — the account, for provenance stamping.
 *
 * We depend on a tiny `MsalLike` interface rather than `@azure/msal-browser` directly, so the
 * core stays unit-testable and free of a heavyweight browser dependency; the app wires a real
 * `IPublicClientApplication` (which satisfies this shape) at startup.
 */
export interface MsalAccount {
  username: string;
  name?: string;
  localAccountId?: string;
  homeAccountId?: string;
}

export interface MsalAuthResult {
  accessToken: string;
  idToken?: string;
  account?: MsalAccount | null;
}

export interface MsalTokenRequest {
  scopes: string[];
  account?: MsalAccount | null;
}

export interface MsalLike {
  getActiveAccount(): MsalAccount | null;
  getAllAccounts?(): MsalAccount[];
  acquireTokenSilent(request: MsalTokenRequest): Promise<MsalAuthResult>;
  /** Optional interactive fallback when silent acquisition needs user interaction. */
  acquireTokenPopup?(request: MsalTokenRequest): Promise<MsalAuthResult>;
}

export interface NaaAuthOptions {
  /** Scopes whose token the Workforce provider trusts as the WIF subject (yields the id token). */
  idTokenScopes: string[];
  /** Default Microsoft Graph scopes used when a caller passes none. */
  graphScopes?: string[];
}

export class NaaAuthClient implements AuthClient, EntraTokenProvider {
  constructor(
    private readonly msal: MsalLike,
    private readonly opts: NaaAuthOptions,
  ) {}

  /** The Entra id token that is the WIF subject. */
  async getIdToken(): Promise<string> {
    const res = await this.acquire(this.opts.idTokenScopes);
    if (!res.idToken) {
      throw new Error('MSAL returned no id token for the WIF subject scopes.');
    }
    return res.idToken;
  }

  /** A delegated Graph access token (falls back to the configured default scopes). */
  async getGraphToken(scopes: string[]): Promise<string> {
    const want = scopes.length ? scopes : (this.opts.graphScopes ?? []);
    const res = await this.acquire(want);
    return res.accessToken;
  }

  async getIdentity(): Promise<UserIdentity> {
    const account = this.account();
    if (!account) throw new Error('No signed-in account to derive identity from.');
    return {
      username: account.username,
      ...(account.name ? { displayName: account.name } : {}),
      ...(account.localAccountId ? { oid: account.localAccountId } : {}),
    };
  }

  private account(): MsalAccount | null {
    return this.msal.getActiveAccount() ?? this.msal.getAllAccounts?.()[0] ?? null;
  }

  private async acquire(scopes: string[]): Promise<MsalAuthResult> {
    const account = this.account();
    try {
      return await this.msal.acquireTokenSilent({ scopes, account });
    } catch (err) {
      if (this.msal.acquireTokenPopup && isInteractionRequired(err)) {
        return this.msal.acquireTokenPopup({ scopes });
      }
      throw err;
    }
  }
}

/** MSAL signals "needs UI" via this error name / a handful of OAuth error codes. */
function isInteractionRequired(err: unknown): boolean {
  const e = err as { name?: string; errorCode?: string } | null;
  if (e?.name === 'InteractionRequiredAuthError') return true;
  const code = e?.errorCode ?? '';
  return (
    code.includes('interaction_required') ||
    code.includes('consent_required') ||
    code.includes('login_required')
  );
}
