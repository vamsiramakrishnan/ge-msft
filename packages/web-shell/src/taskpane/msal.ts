import type { MsalLike, MsalAccount, MsalAuthResult } from '../auth-client.js';
import type { MsalConfigLike } from './config.js';
import type {
  CommonAuthorizationUrlRequest,
  CommonEndSessionRequest,
} from '@azure/msal-common/browser';
import {
  PublicClientApplication,
  createNestablePublicClientApplication,
  type Configuration,
  type IController,
  type IPublicClientApplication,
} from '@azure/msal-browser';

/**
 * Adapt MSAL Nested App Authentication to web-shell's tiny `MsalLike` interface.
 *
 * NAA inside an Office host relies on `createNestablePublicClientApplication`. Prefer a usable
 * host-provided `window.msal` when present, but fall back to the bundled package because Office
 * pages can expose unrelated globals under that name.
 * The shape below is the minimal slice of `IPublicClientApplication` we use; it structurally
 * satisfies `MsalLike`.
 */

interface PublicClientApplicationLike {
  initialize?(): Promise<void>;
  getActiveAccount(): unknown;
  getAllAccounts(): unknown[];
  setActiveAccount?(account: unknown): void;
  ssoSilent?(request: unknown): Promise<unknown>;
  acquireTokenSilent(request: unknown): Promise<unknown>;
  loginPopup?(request: unknown): Promise<unknown>;
  acquireTokenPopup?(request: unknown): Promise<unknown>;
}

interface MsalBrowserGlobal {
  createNestablePublicClientApplication?(config: {
    auth: { clientId: string; authority?: string; redirectUri?: string };
    system?: { popupBridgeTimeout?: number; iframeBridgeTimeout?: number };
  }): Promise<PublicClientApplicationLike>;
  createNestablePublicClientApplicationWithFactory?(
    config: {
      auth: { clientId: string; authority?: string; redirectUri?: string };
      system?: { popupBridgeTimeout?: number; iframeBridgeTimeout?: number };
    },
    correlationId: string | undefined,
    pcaFactory: (configuration: Configuration, controller: IController) => IPublicClientApplication,
  ): Promise<PublicClientApplicationLike>;
  PublicClientApplication?: new (config: {
    auth: { clientId: string; authority?: string; redirectUri?: string };
    system?: { popupBridgeTimeout?: number; iframeBridgeTimeout?: number };
  }) => PublicClientApplicationLike;
}

function importedMsal(): MsalBrowserGlobal {
  return {
    createNestablePublicClientApplication:
      createNestablePublicClientApplication as MsalBrowserGlobal['createNestablePublicClientApplication'],
    createNestablePublicClientApplicationWithFactory:
      createNestablePublicClientApplication as MsalBrowserGlobal['createNestablePublicClientApplicationWithFactory'],
    PublicClientApplication:
      PublicClientApplication as MsalBrowserGlobal['PublicClientApplication'],
  };
}

function msalGlobal(): MsalBrowserGlobal {
  const candidate = (globalThis as { msal?: MsalBrowserGlobal }).msal;
  if (candidate?.createNestablePublicClientApplication || candidate?.PublicClientApplication) {
    return {
      ...candidate,
      ...(candidate.createNestablePublicClientApplication
        ? {
            createNestablePublicClientApplicationWithFactory:
              candidate.createNestablePublicClientApplication as MsalBrowserGlobal['createNestablePublicClientApplicationWithFactory'],
          }
        : {}),
    };
  }
  return importedMsal();
}

interface PopupBridgeMessage {
  type: 'ge-msal-auth-response';
  v: 1;
  id: string;
  payload: string;
}

interface BroadcastBridgeMessage {
  v: 1;
  payload: string;
}

function isPopupBridgeMessage(value: unknown): value is PopupBridgeMessage {
  if (!value || typeof value !== 'object') return false;
  const msg = value as Partial<PopupBridgeMessage>;
  return (
    msg.type === 'ge-msal-auth-response' &&
    msg.v === 1 &&
    typeof msg.id === 'string' &&
    typeof msg.payload === 'string'
  );
}

function isBroadcastBridgeMessage(value: unknown): value is BroadcastBridgeMessage {
  if (!value || typeof value !== 'object') return false;
  const msg = value as Partial<BroadcastBridgeMessage>;
  return msg.v === 1 && typeof msg.payload === 'string';
}

function decodeLibraryStateId(state: string | undefined, win: Window): string | undefined {
  if (!state) return undefined;
  try {
    const encodedLibraryState = state.split('|')[0];
    if (!encodedLibraryState) return undefined;
    const normalized = encodedLibraryState.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, '=');
    const decoded: unknown = JSON.parse(win.atob(padded));
    if (!decoded || typeof decoded !== 'object') return undefined;
    const id = (decoded as { id?: unknown }).id;
    return typeof id === 'string' && id.length > 0 ? id : undefined;
  } catch {
    return undefined;
  }
}

function waitForPopupPostMessage(
  popupWindow: Window,
  popupWindowParent: Window,
  origin: string,
): Promise<string> {
  return new Promise((resolve) => {
    const timeout = popupWindowParent.setTimeout(() => {
      popupWindowParent.removeEventListener('message', onMessage);
    }, 60000);

    function onMessage(event: MessageEvent<unknown>): void {
      if (event.origin !== origin) return;
      if (!isPopupBridgeMessage(event.data)) return;
      popupWindowParent.clearTimeout(timeout);
      popupWindowParent.removeEventListener('message', onMessage);
      resolve(event.data.payload);
    }

    popupWindowParent.addEventListener('message', onMessage);
  });
}

function waitForPopupBroadcastChannel(
  request: CommonAuthorizationUrlRequest | CommonEndSessionRequest,
  popupWindowParent: Window,
): Promise<string> | undefined {
  const channelId = decodeLibraryStateId(
    (request as { state?: string | undefined }).state,
    popupWindowParent,
  );
  const BroadcastChannelCtor = (
    popupWindowParent as Window & { BroadcastChannel?: typeof BroadcastChannel }
  ).BroadcastChannel;
  if (!channelId || typeof BroadcastChannelCtor !== 'function') return undefined;

  return new Promise((resolve) => {
    const channel = new BroadcastChannelCtor(channelId);
    const timeout = popupWindowParent.setTimeout(() => {
      channel.close();
    }, 60000);

    channel.onmessage = (event: MessageEvent<unknown>) => {
      if (!isBroadcastBridgeMessage(event.data)) return;
      popupWindowParent.clearTimeout(timeout);
      channel.close();
      resolve(event.data.payload);
    };
  });
}

class OfficePopupBridgePublicClientApplication extends PublicClientApplication {
  protected override waitForPopupResponse(
    request: CommonAuthorizationUrlRequest | CommonEndSessionRequest,
    popupWindow: Window,
    popupWindowParent: Window,
  ): Promise<string> {
    const broadcastResponse = waitForPopupBroadcastChannel(request, popupWindowParent);
    return Promise.race([
      super.waitForPopupResponse(request, popupWindow, popupWindowParent),
      waitForPopupPostMessage(popupWindow, popupWindowParent, window.location.origin),
      ...(broadcastResponse ? [broadcastResponse] : []),
    ]);
  }
}

function asAccount(value: unknown): MsalAccount | null {
  if (!value || typeof value !== 'object') return null;
  const account = value as Partial<MsalAccount>;
  return typeof account.username === 'string' ? (value as MsalAccount) : null;
}

function asAuthResult(value: unknown): MsalAuthResult {
  if (!value || typeof value !== 'object')
    throw new Error('MSAL returned an invalid token result.');
  const result = value as Partial<MsalAuthResult>;
  if (typeof result.accessToken !== 'string') {
    throw new Error('MSAL returned an invalid token result.');
  }
  return {
    accessToken: result.accessToken,
    ...(typeof result.idToken === 'string' ? { idToken: result.idToken } : {}),
    ...(result.account !== undefined ? { account: asAccount(result.account) } : {}),
  };
}

function adaptMsal(app: PublicClientApplicationLike): MsalLike {
  return {
    getActiveAccount: () => asAccount(app.getActiveAccount()),
    getAllAccounts: () =>
      app
        .getAllAccounts()
        .map(asAccount)
        .filter((a) => a !== null),
    setActiveAccount: (account) => app.setActiveAccount?.(account),
    ssoSilent: app.ssoSilent
      ? async (request) => asAuthResult(await app.ssoSilent?.(request))
      : undefined,
    acquireTokenSilent: async (request) => asAuthResult(await app.acquireTokenSilent(request)),
    loginPopup: app.loginPopup
      ? async (request) => asAuthResult(await app.loginPopup?.(request))
      : undefined,
    acquireTokenPopup: app.acquireTokenPopup
      ? async (request) => asAuthResult(await app.acquireTokenPopup?.(request))
      : undefined,
  };
}

/**
 * Build an `MsalLike` for NAA. Prefers `createNestablePublicClientApplication` (NAA) and falls
 * back to a plain `PublicClientApplication`. Throws a clear error if the host page didn't load
 * MSAL — the bootstrap surfaces that as the unsupported/misconfigured message.
 */
export async function createMsal(cfg: MsalConfigLike): Promise<MsalLike> {
  const msal = msalGlobal();
  const auth = {
    clientId: cfg.clientId,
    ...(cfg.authority ? { authority: cfg.authority } : {}),
    ...(cfg.redirectUri ? { redirectUri: cfg.redirectUri } : {}),
  };
  const system = {
    popupBridgeTimeout: cfg.popupBridgeTimeoutMs ?? 60000,
    iframeBridgeTimeout: cfg.iframeBridgeTimeoutMs ?? 15000,
  };
  const appCandidate = cfg.forceStandardPopupBridge
    ? new OfficePopupBridgePublicClientApplication({ auth, system })
    : msal.createNestablePublicClientApplicationWithFactory
      ? await msal.createNestablePublicClientApplicationWithFactory(
          { auth, system },
          undefined,
          (configuration, controller) =>
            new OfficePopupBridgePublicClientApplication(configuration, controller),
        )
      : msal.createNestablePublicClientApplication
        ? await msal.createNestablePublicClientApplication({ auth, system })
        : msal.PublicClientApplication
          ? new OfficePopupBridgePublicClientApplication({ auth, system })
          : undefined;
  const app = appCandidate as PublicClientApplicationLike | undefined;
  if (!app) {
    throw new Error('MSAL global is present but exposes no usable client application constructor.');
  }
  await app.initialize?.();
  // Ensure an active account so the NaaAuthClient can stamp provenance and silently acquire.
  if (!app.getActiveAccount()) {
    const first = app.getAllAccounts()[0];
    if (first) {
      try {
        app.setActiveAccount?.(first);
      } catch {
        // NAA hosts can expose but reject setActiveAccount; NaaAuthClient also tracks in memory.
      }
    }
  }
  return adaptMsal(app);
}
