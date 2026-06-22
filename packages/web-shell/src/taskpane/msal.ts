import type { MsalLike, MsalAccount, MsalAuthResult, MsalTokenRequest } from '../auth-client.js';
import type { MsalConfigLike } from './config.js';

/**
 * Adapt MSAL Nested App Authentication to web-shell's tiny `MsalLike` interface.
 *
 * `@azure/msal-browser` is intentionally NOT a workspace dependency: NAA inside an Office host
 * relies on `createNestablePublicClientApplication`, and the host page is what version-pins MSAL.
 * So we resolve it at runtime from a global the host page exposes (`window.msal`) rather than
 * importing it — keeping the surface-agnostic core free of a heavyweight, host-pinned browser dep.
 * The shape below is the minimal slice of `IPublicClientApplication` we use; it structurally
 * satisfies `MsalLike`.
 */

interface PublicClientApplicationLike {
  initialize?(): Promise<void>;
  getActiveAccount(): MsalAccount | null;
  getAllAccounts(): MsalAccount[];
  setActiveAccount?(account: MsalAccount | null): void;
  acquireTokenSilent(request: MsalTokenRequest): Promise<MsalAuthResult>;
  acquireTokenPopup?(request: MsalTokenRequest): Promise<MsalAuthResult>;
  ssoSilent?(request: MsalTokenRequest): Promise<MsalAuthResult>;
}

interface MsalBrowserGlobal {
  createNestablePublicClientApplication?(config: {
    auth: { clientId: string; authority?: string };
  }): Promise<PublicClientApplicationLike>;
  PublicClientApplication?: new (config: {
    auth: { clientId: string; authority?: string };
  }) => PublicClientApplicationLike;
}

function msalGlobal(): MsalBrowserGlobal | undefined {
  return (globalThis as { msal?: MsalBrowserGlobal }).msal;
}

/**
 * Build an `MsalLike` for NAA. Prefers `createNestablePublicClientApplication` (NAA) and falls
 * back to a plain `PublicClientApplication`. Throws a clear error if the host page didn't load
 * MSAL — the bootstrap surfaces that as the unsupported/misconfigured message.
 */
export async function createMsal(cfg: MsalConfigLike): Promise<MsalLike> {
  const msal = msalGlobal();
  if (!msal) {
    throw new Error(
      'MSAL is not available. The host page must load @azure/msal-browser (window.msal) so the ' +
        'add-in can federate the signed-in Entra identity via Nested App Authentication.',
    );
  }
  const auth = { clientId: cfg.clientId, ...(cfg.authority ? { authority: cfg.authority } : {}) };
  const app = msal.createNestablePublicClientApplication
    ? await msal.createNestablePublicClientApplication({ auth })
    : msal.PublicClientApplication
      ? new msal.PublicClientApplication({ auth })
      : undefined;
  if (!app) {
    throw new Error('MSAL global is present but exposes no usable client application constructor.');
  }
  await app.initialize?.();
  // Ensure an active account so the NaaAuthClient can stamp provenance and silently acquire.
  if (!app.getActiveAccount()) {
    const first = app.getAllAccounts()[0];
    if (first) app.setActiveAccount?.(first);
  }
  return app;
}
