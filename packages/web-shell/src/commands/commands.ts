import { createApplicationRuntime } from '../runtime-extensions.js';
import {
  createMessageSendHandler,
  activeItemIdResolver,
  type OnSendEvent,
} from '@ge/bridge-outlook';

/**
 * The function-command runtime (`commands.html`). Office function commands and LaunchEvents run
 * here, NOT in the task-pane document — so this file is deliberately tiny and has no React/UI. It
 * registers:
 *   • `openGemini` — a ribbon button action that opens the task pane.
 *   • `summarizeSelection` / `explainSelection` — deterministic right-click actions: read the
 *     host selection, stash a typed seed grounding it as `@this`, then reveal the pane.
 *   • `onMessageSend` — the Outlook `OnMessageSend` (Smart Alerts) gate, pointed at the bridge's
 *     pure on-send handler.
 *
 * The same bundled extension registrations are installed in this separate runtime. A mail-send
 * guard executes here, with a deadline and fail-closed check failure. The built-in bundle has
 * no content-grounding policy; deployments add one explicitly through APPLICATION_EXTENSIONS.
 */

// The seed contract lives in a side-effect-free module so the task pane can read it on boot
// without importing this file (which registers Office.actions on load).
import {
  askSelectionSeedKey,
  ASK_SELECTION_SEED_CHANNEL,
  buildAskSelectionSeed,
  type AskSelectionSeed,
} from './ask-selection-seed.js';
import { surfaceFromHost } from '../host.js';
import type { Surface } from '@ge/contracts';

export { askSelectionSeedKey, buildAskSelectionSeed, type AskSelectionSeed };

interface OfficeActionsLike {
  actions: { associate(id: string, fn: (event: { completed(): void }) => void): void };
  context?: { ui?: { displayDialogAsync?: unknown } };
}

const { triggers: sendRegistry } = createApplicationRuntime();
const handleMessageSend = createMessageSendHandler(sendRegistry, {
  resolveItemId: activeItemIdResolver,
});

/** Ribbon action: open the task pane. The host's `openPage` action does the actual reveal; this
 *  associate keeps a programmatic fallback and completes the command. */
function openGemini(event: { completed(): void }): void {
  event.completed();
}

/** The slice of Office the `askSelection` handler reads — kept minimal + guarded so a host that
 *  exposes only part of it (or none) degrades cleanly rather than throwing. */
interface OfficeSelectionLike {
  context?: {
    document?: {
      getSelectedDataAsync?: (
        coercionType: unknown,
        cb: (result: { status?: unknown; value?: unknown }) => void,
      ) => void;
    };
    host?: unknown;
    mailbox?: unknown;
  };
  CoercionType?: { Text?: unknown };
  /** Office.js 1.13+: reveal the task pane from a function command. */
  addin?: { showAsTaskpane?: () => Promise<unknown> };
}

/** Resolve the active surface so the seed is stashed under the per-surface key (Outlook → mailbox). */
function resolveSurface(office: OfficeSelectionLike | undefined): Surface | undefined {
  const ctx = office?.context;
  if (!ctx) return undefined;
  if (ctx.mailbox) return 'outlook';
  return surfaceFromHost(ctx.host != null ? String(ctx.host) : undefined);
}

/** A persistence sink for the seed — `localStorage` in the host, injectable in tests. */
interface SeedSink {
  setItem(key: string, value: string): void;
}

interface SeedBroadcaster {
  postMessage(message: unknown): void;
  close?(): void;
}

/** Read the active host's text selection. Resolves to `''` when no selection API is available (the
 *  Outlook reading pane and unsupported hosts degrade to an empty, still-valid seed). */
function readSelectedText(office: OfficeSelectionLike | undefined): Promise<string> {
  const doc = office?.context?.document;
  const getSelected = doc?.getSelectedDataAsync;
  if (!getSelected) return Promise.resolve('');
  const coercion = office?.CoercionType?.Text ?? 'text';
  return new Promise<string>((resolve) => {
    try {
      getSelected.call(doc, coercion, (result) => {
        resolve(result?.status === 'failed' ? '' : String(result?.value ?? ''));
      });
    } catch {
      resolve('');
    }
  });
}

/**
 * A right-click action. Reads the host selection, stashes a typed read-only seed (grounding the
 * selection as `@this`) where the task pane picks it up on boot, then reveals the pane. The
 * selection is untrusted host content — it rides as `@this` data, never instructions. Always
 * completes the Office event, even on a read failure, so the command never hangs the host.
 *
 * The task-pane boot (`taskpane/main.tsx`) reads the per-surface {@link askSelectionSeedKey} from
 * `localStorage` once on mount, clears it, validates the version + TTL, and seeds the turn via
 * `controller.send(askSelectionQuery(seed))` — completing the right-click → pane handoff.
 */
async function askSelection(
  event: { completed(): void },
  deps: { office?: OfficeSelectionLike; sink?: SeedSink; broadcaster?: SeedBroadcaster } = {},
  mode: AskSelectionSeed['mode'] = 'ask',
): Promise<void> {
  const office =
    deps.office ?? (globalThis as { Office?: OfficeSelectionLike }).Office ?? undefined;
  const sink = deps.sink ?? (globalThis as { localStorage?: SeedSink }).localStorage ?? undefined;
  try {
    const selection = await readSelectedText(office);
    const seed = buildAskSelectionSeed(selection, Date.now(), mode);
    const surface = resolveSurface(office) ?? 'word';
    try {
      sink?.setItem(askSelectionSeedKey(surface), JSON.stringify(seed));
      publishSeedWritten(surface, deps.broadcaster);
    } catch {
      // A full/blocked storage must not abort the reveal — the pane simply opens without a seed.
    }
    try {
      await office?.addin?.showAsTaskpane?.();
    } catch {
      // The host may reject (e.g. pane already open); ignore — the action still completed.
    }
  } finally {
    event.completed();
  }
}

/** Fixed host-command wrappers keep the manifest action IDs free of arbitrary prompt text. */
function summarizeSelection(event: { completed(): void }): void {
  void askSelection(event, {}, 'summarize');
}

function explainSelection(event: { completed(): void }): void {
  void askSelection(event, {}, 'explain');
}

function publishSeedWritten(surface: Surface, injected?: SeedBroadcaster): void {
  const broadcaster =
    injected ??
    (() => {
      const Channel = (globalThis as { BroadcastChannel?: typeof BroadcastChannel })
        .BroadcastChannel;
      return Channel ? new Channel(ASK_SELECTION_SEED_CHANNEL) : undefined;
    })();
  try {
    broadcaster?.postMessage({ kind: 'ask-selection-seed-written', surface });
  } catch {
    // Best-effort wake-up only; the seed is already in storage for cold-pane boot.
  } finally {
    if (!injected) broadcaster?.close?.();
  }
}

/** The `OnMessageSend` LaunchEvent entry. Delegates to the bridge's fail-safe handler. */
function onMessageSend(event: OnSendEvent): void {
  void handleMessageSend(event);
}

function register(): void {
  const office = (
    globalThis as { Office?: OfficeActionsLike & { onReady?: (cb: () => void) => void } }
  ).Office;
  if (!office) return;
  const associate = (): void => {
    office.actions?.associate('openGemini', openGemini);
    office.actions?.associate('askSelection', (e: { completed(): void }) => void askSelection(e));
    office.actions?.associate('summarizeSelection', summarizeSelection);
    office.actions?.associate('explainSelection', explainSelection);
    office.actions?.associate('onMessageSend', onMessageSend as (e: { completed(): void }) => void);
  };
  if (office.onReady) office.onReady(associate);
  else associate();
}

register();

// Exported for unit testing / explicit host association.
export {
  openGemini,
  askSelection,
  summarizeSelection,
  explainSelection,
  onMessageSend,
  handleMessageSend,
};
