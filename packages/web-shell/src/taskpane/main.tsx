import { StrictMode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  filterManifestForReleaseProfile,
  intentsForManifest,
  isSurfaceEnabledForProfile,
  releaseProfile,
  type Surface,
} from '@ge/contracts';
import {
  applyCatalogSelection,
  defaultCatalogSelection,
  DiscoveryCatalogClient,
  type GeminiCatalog,
  type StreamAssistClient,
} from '@ge/gemini-client';
import { detectSurface, surfaceFromHost } from '../host.js';
import { NaaAuthClient } from '../auth-client.js';
import { composeSession } from '../compose.js';
import { createApplicationRuntime } from '../runtime-extensions.js';
import { connectPanelRuntime } from '../panel-runtime.js';

let disposePanelRuntime: (() => void) | undefined;
import { PanelController } from '../controller.js';
import { App } from './components/App.js';
import { selectBridge, isSupportedSurface } from './select-bridge.js';
import {
  askSelectionSeedKey,
  ASK_SELECTION_SEED_CHANNEL,
  askSelectionQuery,
  isAskSelectionSeed,
  isAskSelectionSeedFresh,
  type AskSelectionSeed,
} from '../commands/ask-selection-seed.js';
import { initialUnit } from './unit.js';
import { createMsal } from './msal.js';
import {
  shellConfigFromEnv,
  authOptionsFromEnv,
  msalConfigFromEnv,
  notebookIdFromEnv,
  releaseProfileFromEnv,
  warmUpSkillsFromEnv,
  type RawEnv,
} from './config.js';
import './styles.css';
import './workspace.css';

/**
 * Task-pane entry point. The Office bootstrap, and the ONLY place the surface-agnostic shell wires
 * a concrete host bridge:
 *
 *   Office.onReady → detect surface → selectBridge(surface) → composeSession(WIF + Discovery
 *   Engine + AssistSession) → new PanelController → render <App/>.
 *
 * Config (project/location/engine/WIF/Entra) is read from `import.meta.env` via `config.ts`; no
 * secret is ever hardcoded here. Failures render a readable message instead of a blank pane.
 */

const env = import.meta.env as unknown as RawEnv;
let root: Root | undefined;

type AuthStart = () => Promise<void>;

interface AuthDebugEvent {
  at: string;
  stage: string;
  detail?: Record<string, string | number | boolean>;
}

const authDebugEvents: AuthDebugEvent[] = [];

function recordAuthDebug(stage: string, detail?: AuthDebugEvent['detail']): void {
  const event: AuthDebugEvent = {
    at: new Date().toISOString(),
    stage,
    ...(detail ? { detail } : {}),
  };
  authDebugEvents.push(event);
  const exposed = authDebugEvents.slice(-50);
  (globalThis as { __GE_AUTH_DEBUG__?: AuthDebugEvent[] }).__GE_AUTH_DEBUG__ = exposed;
  if (env.MODE !== 'production') {
    console.info('[ge-auth]', stage, detail ?? {});
  }
}

interface PreparedBoot {
  surface: Surface;
  profileName?: ReturnType<typeof releaseProfileFromEnv>;
  pendingSeed: AskSelectionSeed | null;
  bridge: NonNullable<ReturnType<typeof selectBridge>>;
  msal: Awaited<ReturnType<typeof createMsal>>;
  nestedAppAuthBridge: boolean;
  loginHint?: string;
}

function mount(node: JSX.Element): void {
  const el = document.getElementById('root');
  if (!el) throw new Error('Missing #root element in taskpane.html');
  root ??= createRoot(el);
  root.render(<StrictMode>{node}</StrictMode>);
}

function fatal(
  title: string,
  detail: string,
  action?: { label: string; onClick: (button: HTMLButtonElement) => void },
): void {
  mount(
    <div className="fatal">
      <h1>{title}</h1>
      <p>{detail}</p>
      {action ? (
        <button
          type="button"
          className="fatal-action"
          onClick={(event) => action.onClick(event.currentTarget)}
        >
          {action.label}
        </button>
      ) : null}
    </div>,
  );
}

/** Detect the surface, falling back to a `?host=` query param (used by the Teams tab). */
function resolveSurface(): Surface | undefined {
  const detected = detectSurface();
  if (detected) return detected;
  const fromQuery = new URLSearchParams(window.location.search).get('host');
  if (fromQuery === 'teams') return 'teams';
  return surfaceFromHost(fromQuery);
}

/**
 * Read + CLEAR any context-menu "Ask Gemini about this" seed left by the `askSelection` command,
 * returning it (validated) or null. Clearing happens unconditionally and up front so a seed never
 * outlives one boot — even if boot later fails — and a foreign/malformed value is rejected rather
 * than trusted. Guarded end-to-end: storage problems just yield null.
 */
function takeAskSelectionSeed(surface: Surface): AskSelectionSeed | null {
  try {
    const key = askSelectionSeedKey(surface);
    const store = (globalThis as { localStorage?: Storage }).localStorage;
    const raw = store?.getItem(key);
    if (!raw) return null;
    store?.removeItem(key); // clear-on-read, before we trust the value
    const parsed: unknown = JSON.parse(raw);
    if (!isAskSelectionSeed(parsed)) return null; // validates kind + version + typed {intent, scope}
    return isAskSelectionSeedFresh(parsed) ? parsed : null; // drop a stale/replayed seed
  } catch {
    return null;
  }
}

function consumeAskSelectionSeed(surface: Surface, controller: PanelController): void {
  const seed = takeAskSelectionSeed(surface);
  if (seed) void controller.send(askSelectionQuery(seed));
}

function listenForAskSelectionSeeds(surface: Surface, controller: PanelController): void {
  const consume = (): void => consumeAskSelectionSeed(surface, controller);
  const storageHandler = (event: StorageEvent): void => {
    if (event.key === askSelectionSeedKey(surface)) consume();
  };
  window.addEventListener('storage', storageHandler);

  try {
    const channel = new BroadcastChannel(ASK_SELECTION_SEED_CHANNEL);
    channel.onmessage = (event: MessageEvent<unknown>) => {
      const message = event.data as { kind?: unknown; surface?: unknown };
      if (message.kind === 'ask-selection-seed-written' && message.surface === surface) consume();
    };
  } catch {
    // Some hosts disable BroadcastChannel; the cold-pane and storage-event paths still work.
  }
}

async function getOfficeLoginHint(): Promise<string | undefined> {
  try {
    const officeAuth = (
      globalThis as {
        Office?: {
          auth?: {
            getAuthContext?: () => Promise<{ userPrincipalName?: string | null }>;
          };
        };
      }
    ).Office?.auth;
    const upn = (await officeAuth?.getAuthContext?.())?.userPrincipalName?.trim();
    return upn || undefined;
  } catch {
    return undefined;
  }
}

function hasMsalAccount(msal: {
  getActiveAccount(): unknown;
  getAllAccounts?(): unknown[];
}): boolean {
  return Boolean(msal.getActiveAccount() ?? msal.getAllAccounts?.()[0]);
}

function hasNestedAppAuthBridge(): boolean {
  return (
    typeof (globalThis as { __initializeNestedAppAuth?: unknown }).__initializeNestedAppAuth ===
    'function'
  );
}

function shouldForceStandardPopupBridge(nestedAppAuthBridge = hasNestedAppAuthBridge()): boolean {
  return env.MODE !== 'production' && !nestedAppAuthBridge;
}

function msalRedirectUri(): string {
  return `${window.location.origin}/auth-redirect.html`;
}

function renderSignInPrompt(startInteractive: AuthStart): void {
  fatal(
    'Sign in required',
    'Use your Microsoft 365 account to start Gemini Enterprise in this Office document.',
    {
      label: 'Sign in',
      onClick: (button) => {
        button.disabled = true;
        button.textContent = 'Signing in...';
        runAuthAttempt(startInteractive, button);
      },
    },
  );
}

function renderSignInRetryPrompt(startInteractiveRetry: AuthStart): void {
  fatal(
    'Sign-in state is stuck',
    'Microsoft still has a pending sign-in marker for this add-in. Reset the temporary sign-in state and try again.',
    {
      label: 'Reset and sign in',
      onClick: (button) => {
        button.disabled = true;
        button.textContent = 'Resetting...';
        recordAuthDebug('ui.resetTemporaryAuthState');
        clearMsalTemporaryAuthState();
        runAuthAttempt(startInteractiveRetry, button);
      },
    },
  );
}

function runAuthAttempt(start: AuthStart, button: HTMLButtonElement): void {
  recordAuthDebug('ui.signInClick');
  const timeout = window.setTimeout(() => {
    recordAuthDebug('ui.signInStillWaiting');
    button.disabled = false;
    button.textContent = 'Reset and sign in';
    renderSignInRetryPrompt(() =>
      boot({ interactiveAuth: true, overrideInteractionInProgress: true }),
    );
  }, 45000);
  start()
    .catch((err: unknown) => {
      recordAuthDebug('ui.signInUnhandledError', summarizeAuthError(err));
      fatal('Could not start Gemini Enterprise', errorMessage(err));
    })
    .finally(() => window.clearTimeout(timeout));
}

function clearMsalTemporaryAuthState(): void {
  const temporarySuffixes = [
    'interaction.status',
    'request.origin',
    'urlHash',
    'request.params',
    'code.verifier',
    'request.native',
  ];
  const shouldRemove = (key: string): boolean =>
    key === 'msal.interaction.status' ||
    (key.startsWith('msal.') &&
      temporarySuffixes.some((suffix) => key === `msal.${suffix}` || key.endsWith(`.${suffix}`)));
  const clearStore = (store: Storage): void => {
    for (const key of Object.keys(store)) {
      if (shouldRemove(key)) store.removeItem(key);
    }
  };
  try {
    clearStore(window.sessionStorage);
  } catch {
    // Storage can be unavailable in some Office web iframe privacy modes; MSAL will surface that.
  }
  try {
    clearStore(window.localStorage);
  } catch {
    // Storage can be unavailable in some Office web iframe privacy modes; MSAL will surface that.
  }
}

interface BootOptions {
  interactiveAuth?: boolean;
  overrideInteractionInProgress?: boolean;
}

async function boot(opts: BootOptions = {}): Promise<void> {
  recordAuthDebug('boot.start', {
    interactiveAuth: opts.interactiveAuth === true,
    overrideInteractionInProgress: opts.overrideInteractionInProgress === true,
  });
  const surface = resolveSurface();
  if (!isSupportedSurface(surface)) {
    recordAuthDebug('boot.unsupportedHost', { surface: surface ?? 'unknown' });
    fatal(
      'Unsupported host',
      surface
        ? `The ${surface} surface is not wired in this shell yet.`
        : 'Could not detect a supported Microsoft 365 host (Word, Excel, PowerPoint, OneNote, Outlook, or Teams).',
    );
    return;
  }
  const profileName = releaseProfileFromEnv(env);
  if (profileName) {
    const profile = releaseProfile(profileName);
    if (!isSurfaceEnabledForProfile(surface, profile)) {
      fatal(
        'Unsupported host',
        `${surface} is disabled by the ${profileName} release profile. This alpha enables Word and Excel only.`,
      );
      return;
    }
  }

  // Take (and clear) the per-surface seed before anything that can throw, so it never persists
  // across a failed boot. The key is namespaced by surface so two open hosts can't cross-read.
  const pendingSeed = takeAskSelectionSeed(surface);

  // The surface seam: one host-specific decision, everything downstream is interface-only.
  const bridge = selectBridge(surface);
  if (!bridge) {
    recordAuthDebug('boot.noBridge', { surface });
    fatal('Unsupported host', `No bridge available for ${surface}.`);
    return;
  }

  try {
    recordAuthDebug('msal.create.start');
    const nestedAppAuthBridge = hasNestedAppAuthBridge();
    const forceStandardPopupBridge = shouldForceStandardPopupBridge(nestedAppAuthBridge);
    const redirectUri = msalRedirectUri();
    recordAuthDebug('msal.bridgeMode', {
      nestedAppAuthBridge,
      forceStandardPopupBridge,
    });
    recordAuthDebug('msal.redirectUri', {
      authRedirect: redirectUri.endsWith('/auth-redirect.html'),
    });
    const msal = await createMsal({
      ...msalConfigFromEnv(env),
      redirectUri,
      forceStandardPopupBridge,
    });
    recordAuthDebug('msal.create.success', { hasAccount: hasMsalAccount(msal) });
    const loginHint = await getOfficeLoginHint();
    if (loginHint) recordAuthDebug('office.loginHint.detected');
    const prepared: PreparedBoot = {
      surface,
      profileName,
      pendingSeed,
      bridge,
      msal: instrumentMsal(msal),
      nestedAppAuthBridge,
      ...(loginHint ? { loginHint } : {}),
    };
    const canTryBrokerSilentSso = nestedAppAuthBridge && Boolean(loginHint);
    if (!opts.interactiveAuth && !hasMsalAccount(msal) && !canTryBrokerSilentSso) {
      recordAuthDebug('boot.needsInteractiveSignIn');
      renderSignInPrompt(() => finishBoot(prepared, { interactiveAuth: true }));
      return;
    }
    await finishBoot(prepared, opts);
  } catch (err) {
    const detail = err instanceof Error ? err.message : String(err);
    recordAuthDebug('boot.error', summarizeAuthError(err));
    if (isInteractionInProgressError(err)) {
      renderSignInRetryPrompt(() =>
        boot({ interactiveAuth: true, overrideInteractionInProgress: true }),
      );
    } else if (!opts.interactiveAuth && isRecoverableAuthBootstrapError(err)) {
      renderSignInPrompt(() => boot({ interactiveAuth: true }));
    } else {
      fatal('Could not start Gemini Enterprise', detail);
    }
  }
}

async function finishBoot(prepared: PreparedBoot, opts: BootOptions = {}): Promise<void> {
  try {
    recordAuthDebug('finishBoot.start', {
      interactiveAuth: opts.interactiveAuth === true,
      overrideInteractionInProgress: opts.overrideInteractionInProgress === true,
    });
    const authOptions = authOptionsFromEnv(env);
    const preferInteractive = opts.interactiveAuth === true && !prepared.nestedAppAuthBridge;
    recordAuthDebug('auth.strategy', {
      nestedAppAuthBridge: prepared.nestedAppAuthBridge,
      preferInteractive,
      hasLoginHint: Boolean(prepared.loginHint),
    });
    const auth = new NaaAuthClient(prepared.msal, {
      ...authOptions,
      ...(prepared.loginHint ? { loginHint: prepared.loginHint } : {}),
      ...(preferInteractive ? { preferInteractive: true } : {}),
      ...(opts.overrideInteractionInProgress ? { overrideInteractionInProgress: true } : {}),
    });
    const config = shellConfigFromEnv(env);
    const unit = initialUnit({
      surface: prepared.surface,
      ...(notebookIdFromEnv(env) ? { notebookId: notebookIdFromEnv(env) } : {}),
    });

    disposePanelRuntime?.();
    const runtime = createApplicationRuntime();
    disposePanelRuntime = runtime.dispose;
    recordAuthDebug('compose.start');
    const { session, tokens, client, warmUp, availableAgents, availableDataStores } =
      await composeSession({
        config,
        auth,
        bridge: prepared.bridge,
        unit,
        hooks: runtime.hooks,
        triggers: runtime.triggers,
        primeOnHostEvent: false,
        // Self-provision our skills as the signed-in user (client-direct, ADR-0001). Detect-only
        // from env alone (compares the [rev:<sha>] marker); best-effort — never blocks the session.
        warmUpSkills: warmUpSkillsFromEnv(env),
        ...(config.wif.userProject ? { warmUpQuotaProject: config.wif.userProject } : {}),
        // Discover skills (:listAvailableAgentViews) + federated data stores (engines.get) with the
        // WIF token — the sources for the skill / data-store pickers.
        discoverCatalog: true,
      });
    recordAuthDebug('compose.success');
    if (warmUp.length)
      recordAuthDebug('warmup.skills', { results: warmUp.map((w) => w.action).join(',') });
    recordAuthDebug('catalog.discovered', {
      agents: availableAgents.length,
      dataStores: availableDataStores.length,
      connectors: [...new Set(availableDataStores.map((d) => d.connector))].join(','),
    });
    const catalogClient = new DiscoveryCatalogClient(tokens, config);
    await applyBootCatalogRouting(client, catalogClient);
    const controller = new PanelController(session, prepared.bridge);
    controller.setDiscoveredCatalog(availableAgents, availableDataStores);
    const panelRuntime = connectPanelRuntime({
      session,
      bridge: prepared.bridge,
      controller,
      triggers: runtime.triggers,
    });
    const suspend = (): void => {
      panelRuntime.stop();
      controller.cancel();
    };
    const resume = (): void => panelRuntime.start();
    window.addEventListener('pagehide', suspend);
    window.addEventListener('pageshow', resume);
    disposePanelRuntime = () => {
      window.removeEventListener('pagehide', suspend);
      window.removeEventListener('pageshow', resume);
      panelRuntime.dispose();
      controller.cancel();
      session.dispose();
      runtime.dispose();
    };
    panelRuntime.start();

    // Narrow the quick-action bar + `/` palette to what THIS surface can actually run (ADR-0006).
    const rawCapabilities = await prepared.bridge.getCapabilities();
    const capabilities = prepared.profileName
      ? filterManifestForReleaseProfile(rawCapabilities, prepared.profileName)
      : rawCapabilities;
    const allowedIntents = intentsForManifest(capabilities);

    mount(
      <App
        controller={controller}
        surface={prepared.surface}
        allowedIntents={allowedIntents}
        catalogClient={catalogClient}
        onCatalogRouting={(routing) => client.configureRouting(routing)}
      />,
    );
    listenForAskSelectionSeeds(prepared.surface, controller);
    recordAuthDebug('app.mounted');
    // The selection was re-grounded as @this by the bridge; the query is a fixed template.
    if (prepared.pendingSeed) void controller.send(askSelectionQuery(prepared.pendingSeed));
  } catch (err) {
    disposePanelRuntime?.();
    disposePanelRuntime = undefined;
    const detail = err instanceof Error ? err.message : String(err);
    recordAuthDebug('finishBoot.error', summarizeAuthError(err));
    if (isInteractionInProgressError(err)) {
      renderSignInRetryPrompt(() =>
        finishBoot(prepared, { interactiveAuth: true, overrideInteractionInProgress: true }),
      );
    } else if (!opts.interactiveAuth && isRecoverableAuthBootstrapError(err)) {
      renderSignInPrompt(() => finishBoot(prepared, { interactiveAuth: true }));
    } else {
      fatal('Could not start Gemini Enterprise', detail);
    }
  }
}

async function applyBootCatalogRouting(
  client: StreamAssistClient,
  catalogClient: DiscoveryCatalogClient,
): Promise<void> {
  const controller = new AbortController();
  const timeout = window.setTimeout(() => controller.abort(), 3500);
  try {
    const catalog = await catalogClient.listCatalog(controller.signal);
    client.configureRouting(applyCatalogSelection(defaultCatalogSelection(catalog)));
    recordAuthDebug('catalog.routing.discovered', summarizeCatalog(catalog));
  } catch (err) {
    recordAuthDebug('catalog.routing.fallback', {
      error: err instanceof Error ? err.message.slice(0, 240) : String(err).slice(0, 240),
    });
  } finally {
    window.clearTimeout(timeout);
  }
}

function summarizeCatalog(catalog: GeminiCatalog): Record<string, string | number | boolean> {
  const defaults = defaultCatalogSelection(catalog);
  return {
    skills: catalog.skills.length,
    dataStores: catalog.dataStores.length,
    connectors: catalog.connectors.length,
    warnings: catalog.warnings?.length ?? 0,
    planner: defaults.plannerSkill?.label ?? '',
    commander: defaults.commandSkill?.label ?? '',
  };
}

function instrumentMsal(msal: PreparedBoot['msal']): PreparedBoot['msal'] {
  return {
    ...msal,
    ssoSilent: msal.ssoSilent
      ? async (request) => {
          recordAuthDebug('msal.ssoSilent.start', summarizeMsalRequest(request));
          try {
            const result = await msal.ssoSilent!(request);
            recordAuthDebug('msal.ssoSilent.success', summarizeMsalResult(result));
            return result;
          } catch (err) {
            recordAuthDebug('msal.ssoSilent.error', summarizeAuthError(err));
            throw err;
          }
        }
      : undefined,
    acquireTokenSilent: async (request) => {
      recordAuthDebug('msal.acquireTokenSilent.start', summarizeMsalRequest(request));
      try {
        const result = await msal.acquireTokenSilent(request);
        recordAuthDebug('msal.acquireTokenSilent.success', summarizeMsalResult(result));
        return result;
      } catch (err) {
        recordAuthDebug('msal.acquireTokenSilent.error', summarizeAuthError(err));
        throw err;
      }
    },
    loginPopup: msal.loginPopup
      ? async (request) => {
          recordAuthDebug('msal.loginPopup.start', summarizeMsalRequest(request));
          try {
            const result = await msal.loginPopup!(request);
            recordAuthDebug('msal.loginPopup.success', summarizeMsalResult(result));
            return result;
          } catch (err) {
            recordAuthDebug('msal.loginPopup.error', summarizeAuthError(err));
            throw err;
          }
        }
      : undefined,
    acquireTokenPopup: msal.acquireTokenPopup
      ? async (request) => {
          recordAuthDebug('msal.acquireTokenPopup.start', summarizeMsalRequest(request));
          try {
            const result = await msal.acquireTokenPopup!(request);
            recordAuthDebug('msal.acquireTokenPopup.success', summarizeMsalResult(result));
            return result;
          } catch (err) {
            recordAuthDebug('msal.acquireTokenPopup.error', summarizeAuthError(err));
            throw err;
          }
        }
      : undefined,
  };
}

function summarizeMsalRequest(request: {
  scopes?: string[];
  account?: unknown;
  loginHint?: string;
  prompt?: string;
  overrideInteractionInProgress?: boolean;
}): Record<string, string | number | boolean> {
  return {
    scopes: request.scopes?.join(' ') ?? '',
    hasAccount: request.account !== undefined,
    hasLoginHint: Boolean(request.loginHint),
    ...(request.prompt ? { prompt: request.prompt } : {}),
    overrideInteractionInProgress: request.overrideInteractionInProgress === true,
  };
}

function summarizeMsalResult(result: {
  accessToken?: string;
  idToken?: string;
  account?: { username?: string } | null;
}): Record<string, string | number | boolean> {
  return {
    hasAccessToken: typeof result.accessToken === 'string' && result.accessToken.length > 0,
    hasIdToken: typeof result.idToken === 'string' && result.idToken.length > 0,
    hasAccount: Boolean(result.account),
  };
}

function summarizeAuthError(err: unknown): Record<string, string | number | boolean> {
  const e = err as {
    name?: string;
    errorCode?: string;
    subError?: string;
    message?: string;
    status?: string;
  } | null;
  return {
    name: e?.name ?? '',
    errorCode: e?.errorCode ?? '',
    subError: e?.subError ?? '',
    status: e?.status ?? '',
    message: (e?.message ?? String(err)).slice(0, 240),
  };
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function isRecoverableAuthBootstrapError(err: unknown): boolean {
  const e = err as {
    name?: string;
    errorCode?: string;
    subError?: string;
    message?: string;
  } | null;
  const text = `${e?.name ?? ''} ${e?.errorCode ?? ''} ${e?.subError ?? ''} ${e?.message ?? ''}`;
  return /interaction_required|consent_required|login_required|no_account|timed_out|redirect_bridge_timeout|popup/i.test(
    text,
  );
}

function isInteractionInProgressError(err: unknown): boolean {
  const e = err as { errorCode?: string; message?: string } | null;
  return /interaction_in_progress/i.test(`${e?.errorCode ?? ''} ${e?.message ?? ''}`);
}

// Office.js may be absent when the page is opened outside a host (e.g. the Teams tab in a browser);
// fall back to booting directly so `?host=teams` still works.
const office = (globalThis as { Office?: { onReady?: (cb: () => void) => void } }).Office;
if (office?.onReady) {
  office.onReady(() => void boot());
} else {
  void boot();
}
