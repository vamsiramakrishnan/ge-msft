import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import {
  mailItemToContext,
  mailItemToDocStateBlocks,
  searchMailItem,
  type MailItem,
} from './capture.js';
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
/**
 * The exact `ActuationKind`s {@link OutlookBridge.actuate} handles (ADR-0006 closure source of
 * truth). `create-mail` is deliberately ABSENT — it was advertised but never handled, so it was
 * un-advertised. The conformance test asserts this equals the manifest's advertised kinds.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = ['reply-mail'];

export class OutlookBridge implements DocBridge {
  readonly surface = 'outlook' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

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
    const mail = await this.readMailItem();
    return mail ? mailItemToContext(mail) : [];
  }

  /**
   * ADR-0006 whole-item `read` (the runtime's empty-selector read → `captureDocState`): a mail item
   * has no addressable sub-range, so the "document" is the single active item. Builds a snapshot
   * from the subject (heading) + sender + leading body lines (bounded). Read-only; no active item →
   * `undefined`. Version increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    const mail = await this.readMailItem();
    if (!mail) return undefined;
    const blocks = mailItemToDocStateBlocks(mail);
    if (blocks.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'outlook',
      version: this.docStateVersion,
      ...(mail.subject?.trim() ? { title: mail.subject } : {}),
      blocks,
    });
  }

  /**
   * ADR-0006 `search` read: scan the active mail item's body for `query` and return matching lines
   * as `ResolvedContext` data (never instructions), bounded by `searchMailItem`. Scoped strictly to
   * the active item (no cross-mailbox read). Empty query / no active item / no match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q) return [];
    const mail = await this.readMailItem();
    return mail ? searchMailItem(mail, q) : [];
  }

  /**
   * Read the **active** mail item (subject / from / HTML body) into a pure {@link MailItem} — the
   * shared host read behind `resolveContext`/`captureDocState`/`searchDocument`. Read-only and
   * scoped to `Office.context.mailbox.item` (the one open item); never enumerates the mailbox. No
   * active item → `undefined`.
   */
  private async readMailItem(): Promise<MailItem | undefined> {
    const item = Office.context.mailbox?.item;
    if (!item) return undefined;
    const body = await getAsync<string>((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
    return {
      ...(item.itemId ? { id: item.itemId } : {}),
      ...(readSubject(item) ? { subject: readSubject(item) } : {}),
      ...(readFrom(item) ? { from: readFrom(item) } : {}),
      body,
      bodyType: 'html',
    };
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
