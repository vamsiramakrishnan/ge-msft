import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { intentsForManifest, type Surface } from '@ge/contracts';
import { detectSurface, surfaceFromHost } from '../host.js';
import { NaaAuthClient } from '../auth-client.js';
import { composeSession } from '../compose.js';
import { PanelController } from '../controller.js';
import { App } from './components/App.js';
import { selectBridge, isSupportedSurface } from './select-bridge.js';
import {
  askSelectionSeedKey,
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
  type RawEnv,
} from './config.js';
import './styles.css';

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

function mount(node: JSX.Element): void {
  const el = document.getElementById('root');
  if (!el) throw new Error('Missing #root element in taskpane.html');
  createRoot(el).render(<StrictMode>{node}</StrictMode>);
}

function fatal(title: string, detail: string): void {
  mount(
    <div className="fatal">
      <h1>{title}</h1>
      <p>{detail}</p>
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

async function boot(): Promise<void> {
  const surface = resolveSurface();
  if (!isSupportedSurface(surface)) {
    fatal(
      'Unsupported host',
      surface
        ? `The ${surface} surface is not wired in this shell yet.`
        : 'Could not detect a supported Microsoft 365 host (Word, Excel, PowerPoint, OneNote, Outlook, or Teams).',
    );
    return;
  }

  // Take (and clear) the per-surface seed before anything that can throw, so it never persists
  // across a failed boot. The key is namespaced by surface so two open hosts can't cross-read.
  const pendingSeed = takeAskSelectionSeed(surface);

  // The surface seam: one host-specific decision, everything downstream is interface-only.
  const bridge = selectBridge(surface);
  if (!bridge) {
    fatal('Unsupported host', `No bridge available for ${surface}.`);
    return;
  }

  try {
    const msal = await createMsal(msalConfigFromEnv(env));
    const auth = new NaaAuthClient(msal, authOptionsFromEnv(env));
    const config = shellConfigFromEnv(env);
    const unit = initialUnit({
      surface,
      ...(notebookIdFromEnv(env) ? { notebookId: notebookIdFromEnv(env) } : {}),
    });

    const { session } = await composeSession({ config, auth, bridge, unit });
    const controller = new PanelController(session, bridge);

    // Narrow the quick-action bar + `/` palette to what THIS surface can actually run (ADR-0006).
    const allowedIntents = intentsForManifest(await bridge.getCapabilities());

    mount(<App controller={controller} surface={surface} allowedIntents={allowedIntents} />);
    // The selection was re-grounded as @this by the bridge; the query is a fixed template.
    if (pendingSeed) void controller.send(askSelectionQuery(pendingSeed));
  } catch (err) {
    fatal('Could not start Gemini Enterprise', err instanceof Error ? err.message : String(err));
  }
}

// Office.js may be absent when the page is opened outside a host (e.g. the Teams tab in a browser);
// fall back to booting directly so `?host=teams` still works.
const office = (globalThis as { Office?: { onReady?: (cb: () => void) => void } }).Office;
if (office?.onReady) {
  office.onReady(() => void boot());
} else {
  void boot();
}
