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
 * FAIL SAFE (decision path only): if computing the decision throws, we call
 * `completed({ allowEvent: true })`. A guard add-in that crashes while deciding must never wedge
 * the user's Send button — refusing on our own bug is worse than letting the mail through. But the
 * fail-safe is scoped to the DECISION: once a `block` is decided, a crash inside `event.completed`
 * must NOT be downgraded to an allow (that would silently send a blocked mail). The strict decision
 * stays in `decideSend`; only the decision-computation step of this host boundary swallows errors.
 */
export function createMessageSendHandler(
  registry: TriggerRegistry,
  opts: MessageSendHandlerOptions = {},
): MessageSendHandler {
  return async (event: OnSendEvent): Promise<void> => {
    // Two distinct failure domains. (1) Computing the decision: if THAT throws we fail safe and
    // let Send proceed — refusing on our own bug is worse than the mail going out. (2) Signalling
    // the host via `event.completed`: once a decision is reached we must NOT re-call `completed`
    // with `allowEvent: true`, or a real block would silently downgrade to an allow if the first
    // `completed(...)` throws. The `decided` flag separates the two: the catch only fails open
    // while no decision had been committed yet.
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
      // Fail safe ONLY for errors while deciding. If a decision was already committed, a crash
      // happened inside `event.completed`; re-allowing would downgrade a genuine block — re-throw.
      if (decided) throw err instanceof Error ? err : new Error(String(err));
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
