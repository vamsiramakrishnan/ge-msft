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
 * Required checks fail closed. Decision-path failures complete once with a recoverable block;
 * failure to signal a decided result propagates and never triggers a second, contradictory allow.
 */
export function createMessageSendHandler(
  registry: TriggerRegistry,
  opts: MessageSendHandlerOptions = {},
): MessageSendHandler {
  return async (event: OnSendEvent): Promise<void> => {
    // Computing and signalling a decision are separate failure domains. Never downgrade an
    // established block if the host's completion callback throws.
    let decided = false;
    try {
      const id = opts.resolveItemId?.();
      const outcome = await registry.gate(sendEvent(id));
      const decision = decideSend(outcome);
      // Decision computed — from here on a throw must propagate, never fail open.
      decided = true;
      if (!decision.allowEvent) {
        event.completed(
          decision.message !== undefined
            ? { allowEvent: false, errorMessage: decision.message }
            : { allowEvent: false },
        );
        return;
      }
      event.completed({ allowEvent: true });
    } catch (err) {
      if (decided) throw err instanceof Error ? err : new Error(String(err));
      event.completed({
        allowEvent: false,
        errorMessage: 'The send checks could not complete. Try again before sending this message.',
      });
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
