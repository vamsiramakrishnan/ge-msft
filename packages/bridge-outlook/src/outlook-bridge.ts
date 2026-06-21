import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import { mailItemToContext, type MailItem } from './capture.js';
import { planReply } from './actuate-plan.js';
import { composeEvent, receivedEvent } from './events.js';

/**
 * The Outlook `DocBridge`. The ONLY place Office.js mailbox APIs are touched. Unlike Word/Excel
 * (`Word.run` / `Excel.run`), Outlook exposes a single `Office.context.mailbox.item` whose reads
 * are **callback-based** (`getAsync`) — so a small `promisify` helper wraps them into awaitables.
 * Context capture is the string path (the body is HTML/text, not a native object model): reads
 * subject/from/body and hands them to `mailItemToContext` (pure, unit-tested). Writes open a
 * reviewable host form (`displayReplyForm`) rather than mutating silently. Pure mapping lives in
 * `capture.ts` / `actuate-plan.ts`; this file is the host wiring.
 */
export class OutlookBridge implements DocBridge {
  readonly surface = 'outlook' as const;

  getCapabilities(): CapabilityManifest {
    return OUTLOOK_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    const item = Office.context.mailbox?.item;
    if (!item) return [];
    const subject = readSubject(item);
    return [
      {
        id: item.itemId ?? 'outlook:item',
        kind: 'mail-item',
        surface: 'outlook',
        title: subject ? `Email: ${subject}` : 'Current email',
        ...(subject ? { preview: subject.slice(0, 120) } : {}),
        live: true,
      },
    ];
  }

  async resolveContext(_ref: ContextRef): Promise<ResolvedContext[]> {
    const item = Office.context.mailbox?.item;
    if (!item) return [];
    const body = await getAsync<string>((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
    const mail: MailItem = {
      ...(item.itemId ? { id: item.itemId } : {}),
      ...(readSubject(item) ? { subject: readSubject(item) } : {}),
      ...(readFrom(item) ? { from: readFrom(item) } : {}),
      body,
      bodyType: 'html',
    };
    return mailItemToContext(mail);
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    switch (req.kind) {
      case 'reply-mail':
        return this.applyReply(req);
      default:
        return {
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'unsupported', message: `Outlook bridge cannot ${req.kind}` },
        };
    }
  }

  /**
   * Stream Outlook host activity into the trigger engine. From the task pane the one observable
   * lifecycle event is `ItemChanged` (Mailbox 1.5): it fires when the user switches the active
   * item while the pane is pinned. We map the new active item to `mail-compose` (a draft, no
   * `itemId`) or `mail-received` (a saved item) and emit it.
   *
   * Launch events (`OnNewMessageCompose` / `OnMessageSend`) are manifest-declared and run in the
   * separate function-file runtime — they are NOT registered here. The on-send gate is exposed as
   * a factory in `on-send.ts` instead.
   *
   * Defensive throughout: the mailbox may be undefined (wrong host / pane not in a mailbox), and
   * `watch` must never throw — a failed registration just yields a no-op unsubscribe.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    const mailbox = Office.context.mailbox;
    if (!mailbox) return () => {};

    const handler = (): void => {
      try {
        const event = activeItemEvent(mailbox.item);
        if (event) emit(event);
      } catch {
        // An event-source failure must not break the host; drop this notification.
      }
    };

    try {
      mailbox.addHandlerAsync(Office.EventType.ItemChanged, handler);
    } catch {
      return () => {};
    }

    return () => {
      try {
        mailbox.removeHandlerAsync(Office.EventType.ItemChanged, handler);
      } catch {
        // Best-effort teardown; ignore if the host already tore the handler down.
      }
    };
  }

  private async applyReply(req: ActuationRequest): Promise<ActuationResult> {
    const item = Office.context.mailbox?.item;
    if (!item) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_item', message: 'No active mail item to reply to.' },
      };
    }
    const plan = planReply(req);
    if (!plan.body.trim()) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_body', message: 'reply-mail needs params.mail.body or params.text' },
      };
    }
    // Open a reviewable reply form rather than sending — the user confirms before it leaves.
    item.displayReplyForm({ htmlBody: plan.body });
    return { ok: true, changeId: req.changeId, kind: req.kind, location: 'reply-form' };
  }
}

/**
 * Classify the newly active mail item into a HostEvent. Read-mode items expose a stable
 * `itemId` → `mail-received`; compose-mode drafts have no saved id → `mail-compose`. Returns
 * undefined when no item is active (nothing to emit). Reads only the opaque `itemId`, never the
 * subject/body (untrusted content stays out of the event).
 */
function activeItemEvent(item: Office.Item | undefined): HostEvent | undefined {
  if (!item) return undefined;
  const id = readItemId(item);
  return id ? receivedEvent(id) : composeEvent();
}

/** Read the saved item id (read mode) as a non-empty string; undefined for unsaved drafts. */
function readItemId(item: Office.Item): string | undefined {
  const id: unknown = (item as { itemId?: unknown }).itemId;
  return typeof id === 'string' && id.length > 0 ? id : undefined;
}

/** Read the subject as a string (read-mode shape); empty when unavailable. */
function readSubject(item: Office.Item): string | undefined {
  const subject: unknown = (item as { subject?: unknown }).subject;
  return typeof subject === 'string' && subject.length > 0 ? subject : undefined;
}

/** Read the sender's SMTP address (read-mode `EmailAddressDetails`); undefined when unavailable. */
function readFrom(item: Office.Item): string | undefined {
  const from: unknown = (item as { from?: unknown }).from;
  if (from && typeof from === 'object' && 'emailAddress' in from) {
    const email = (from as { emailAddress?: unknown }).emailAddress;
    if (typeof email === 'string' && email.length > 0) return email;
  }
  return undefined;
}

/**
 * Promisify Outlook's callback-based `getAsync` shape: invoke the host call, resolve on
 * `Succeeded`, reject on failure. Keeps the bridge body await-flat like `Word.run`.
 */
function getAsync<T>(
  call: (callback: (result: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    call((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error?.message ?? 'Outlook getAsync failed'));
      }
    });
  });
}
