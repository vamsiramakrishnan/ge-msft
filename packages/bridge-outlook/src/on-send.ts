import type { TriggerRegistry } from '@ge/triggers';
import { decideSend, sendEvent } from './events.js';

/**
 * Host glue for the Outlook on-send gate (the "PreToolUse for email" analog). This is the only
 * Outlook-specific wiring for `OnMessageSend`; the *decision* lives in the pure `decideSend`.
 *
 * NOTE: a real `OnMessageSend` LaunchEvent runs in the function-file runtime and is handed a
 * Smart Alerts event whose `completed` accepts `{ allowEvent, errorMessage }`
 * (`Office.MailboxEvent` / `SmartAlertsEventCompletedOptions`). `Office.AddinCommands.Event`
 * shares the `allowEvent` field. We type the parameter structurally over the shared surface so
 * the handler is assignable from the host event AND trivially stubbable in tests.
 */
export interface OnSendCompletedOptions {
  allowEvent?: boolean;
  /** Smart Alerts dialog message shown when the send is blocked (Markdown-capable in the host). */
  errorMessage?: string;
}

/** The minimal shape of the event object passed to an `OnMessageSend` handler. */
export interface OnSendEvent {
  completed(options?: OnSendCompletedOptions): void;
}

export interface MessageSendHandlerOptions {
  /** Resolve the active draft's item id for the `mail-send` event. Defensive: may return undefined. */
  resolveItemId?: () => string | undefined;
}

export type MessageSendHandler = (event: OnSendEvent) => Promise<void>;

/**
 * Build the `OnMessageSend` handler. It constructs a `mail-send` HostEvent, runs it through the
 * trigger registry as a veto gate, maps the outcome with the pure `decideSend`, and signals the
 * host via `event.completed`.
 *
 * FAIL SAFE: if anything in the gate path throws, we still call `completed({ allowEvent: true })`.
 * A guard add-in that crashes must never permanently wedge the user's Send button — refusing to
 * send on our own bug would be worse than letting the mail through. The strict decision stays in
 * `decideSend`; only this host boundary swallows errors.
 */
export function createMessageSendHandler(
  registry: TriggerRegistry,
  opts: MessageSendHandlerOptions = {},
): MessageSendHandler {
  return async (event: OnSendEvent): Promise<void> => {
    try {
      const id = opts.resolveItemId?.();
      const outcome = await registry.gate(sendEvent(id));
      const decision = decideSend(outcome);
      if (!decision.allowEvent) {
        event.completed(
          decision.message !== undefined
            ? { allowEvent: false, errorMessage: decision.message }
            : { allowEvent: false },
        );
        return;
      }
      event.completed({ allowEvent: true });
    } catch {
      // Fail safe: never block Send on our own error. See the doc comment above.
      event.completed({ allowEvent: true });
    }
  };
}

/** Resolve the active draft's saved item id (undefined for unsaved drafts), defensively. */
export function activeItemIdResolver(): string | undefined {
  try {
    const item: unknown = Office.context.mailbox?.item;
    if (item && typeof item === 'object' && 'itemId' in item) {
      const id = (item as { itemId?: unknown }).itemId;
      if (typeof id === 'string' && id.length > 0) return id;
    }
  } catch {
    // Office unavailable / not in a mailbox context.
  }
  return undefined;
}
