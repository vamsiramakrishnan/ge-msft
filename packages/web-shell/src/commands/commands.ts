import { TriggerRegistry } from '@ge/triggers';
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
 *   • `onMessageSend` — the Outlook `OnMessageSend` (Smart Alerts) gate, pointed at the bridge's
 *     pure on-send handler.
 *
 * NOTE: a fully-grounded send gate composes a session + trigger set; here we wire the structural
 * seam with an empty `TriggerRegistry` (which fails safe → allows) so the LaunchEvent is real and
 * the deploy can grow gating triggers without touching the manifest.
 */

interface OfficeActionsLike {
  actions: { associate(id: string, fn: (event: { completed(): void }) => void): void };
  context?: { ui?: { displayDialogAsync?: unknown } };
}

const sendRegistry = new TriggerRegistry();
const handleMessageSend = createMessageSendHandler(sendRegistry, {
  resolveItemId: activeItemIdResolver,
});

/** Ribbon action: open the task pane. The host's `openPage` action does the actual reveal; this
 *  associate keeps a programmatic fallback and completes the command. */
function openGemini(event: { completed(): void }): void {
  event.completed();
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
    office.actions?.associate('onMessageSend', onMessageSend as (e: { completed(): void }) => void);
  };
  if (office.onReady) office.onReady(associate);
  else associate();
}

register();

// Exported for unit testing / explicit host association.
export { openGemini, onMessageSend, handleMessageSend };
