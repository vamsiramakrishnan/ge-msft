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
import {
  planAttachment,
  planCompose,
  planDraftBody,
  planRecipients,
  planReply,
} from './actuate-plan.js';
import { composeEvent, receivedEvent } from './events.js';

/**
 * The Outlook `DocBridge`. The ONLY place Office.js mailbox APIs are touched. Unlike Word/Excel
 * (`Word.run` / `Excel.run`), Outlook exposes a single `Office.context.mailbox.item` whose reads
 * are **callback-based** (`getAsync`) — so a small `promisify` helper wraps them into awaitables.
 * Context capture is the string path (the body is HTML/text, not a native object model): reads
 * subject/from/body and hands them to `mailItemToContext` (pure, unit-tested). Writes open a
 * reviewable host form (`displayReplyForm`) or edit the OPEN DRAFT in place (compose mode) —
 * nothing ever sends. Pure mapping lives in `capture.ts` / `actuate-plan.ts`; this file is the
 * host wiring.
 */
/**
 * The exact `ActuationKind`s {@link OutlookBridge.actuate} handles (ADR-0006 closure source of
 * truth). `reply-mail` opens a reviewable reply form; `create-mail` opens a fresh draft via
 * `displayNewMessageForm`; `set-recipients` / `add-attachment` / `set-body` / `set-subject` edit
 * the OPEN DRAFT in place (compose mode) — all reviewable, none ever sends. The conformance test
 * asserts this equals the manifest's advertised kinds.
 */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = [
  'reply-mail',
  'create-mail',
  'set-recipients',
  'add-attachment',
  'set-body',
  'set-subject',
];

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

  canRevealContext(ref: ContextRef): boolean {
    return (
      ref.surface === 'outlook' &&
      ref.kind === 'mail-item' &&
      outlookItemIdFromRef(ref) !== undefined
    );
  }

  async revealContext(ref: ContextRef): Promise<void> {
    const itemId = outlookItemIdFromRef(ref);
    if (!itemId) return;
    const mailbox = Office.context.mailbox;
    if (!mailbox) return;
    if (typeof mailbox.displayMessageForm === 'function') {
      mailbox.displayMessageForm(itemId);
      return;
    }
    if (typeof mailbox.displayMessageFormAsync === 'function') {
      await getAsync<void>((cb) => mailbox.displayMessageFormAsync(itemId, cb));
    }
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
      case 'create-mail':
        return this.applyCompose(req);
      case 'set-recipients':
        return this.applyRecipients(req);
      case 'add-attachment':
        return this.applyAttachment(req);
      case 'set-body':
        return this.applyBody(req);
      case 'set-subject':
        return this.applySubject(req);
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

  /**
   * `create-mail`: open a brand-new draft via `displayNewMessageForm` — a reviewable host form the
   * user edits and sends, never an auto-send. Recipients are passed through only when the agent
   * supplied them (by default the draft is left unaddressed). Requires the mailbox host and a
   * non-empty subject; otherwise a corrective so the model can self-fix.
   */
  private async applyCompose(req: ActuationRequest): Promise<ActuationResult> {
    const mailbox = Office.context.mailbox;
    if (!mailbox?.displayNewMessageForm) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_mailbox', message: 'Compose is unavailable on this host.' },
      };
    }
    const plan = planCompose(req);
    if (!plan.subject.trim()) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_subject', message: 'create-mail needs params.mail.subject' },
      };
    }
    mailbox.displayNewMessageForm({
      subject: plan.subject,
      htmlBody: plan.body,
      ...(plan.to ? { toRecipients: plan.to } : {}),
    });
    return { ok: true, changeId: req.changeId, kind: req.kind, location: 'new-message-form' };
  }

  /**
   * Resolve the ACTIVE item into a compose-mode draft (the only target of the in-place draft
   * edits). No active item → `no_item`; an item without the compose write callbacks (a read-mode
   * message) → `no_compose`. Both are correctives so the model can self-fix.
   */
  private resolveDraft(
    req: ActuationRequest,
  ): { draft: DraftItem; error?: undefined } | { draft?: undefined; error: ActuationResult } {
    const item = Office.context.mailbox?.item;
    if (!item) return { error: actuationError(req, 'no_item', 'No active mail item.') };
    const draft = draftView(item);
    return draft
      ? { draft }
      : {
          error: actuationError(
            req,
            'no_compose',
            'The active mail item is not an open draft (compose mode).',
          ),
        };
  }

  /**
   * `set-recipients`: write To/Cc/Bcc on the OPEN DRAFT via `Recipients.setAsync` (replace) or
   * `.addAsync` (append). All requested fields are validated up front so a missing field never
   * leaves a partially-applied write. Send control stays entirely with the user.
   */
  private async applyRecipients(req: ActuationRequest): Promise<ActuationResult> {
    const resolved = this.resolveDraft(req);
    if (resolved.error) return resolved.error;
    const plan = planRecipients(req);
    if (plan.fields.length === 0) {
      return actuationError(
        req,
        'no_recipients',
        'set-recipients needs params.mail.to, .cc, or .bcc',
      );
    }
    const targets: Array<{ recipients: DraftRecipients; addresses: string[] }> = [];
    for (const { field, addresses } of plan.fields) {
      const recipients = draftRecipients(resolved.draft, field);
      if (!recipients) {
        return actuationError(
          req,
          'no_compose',
          `${field} recipients are unavailable on this draft.`,
        );
      }
      targets.push({ recipients, addresses });
    }
    try {
      for (const { recipients, addresses } of targets) {
        await getAsync<void>((cb) =>
          plan.mode === 'add'
            ? recipients.addAsync(addresses, cb)
            : recipients.setAsync(addresses, cb),
        );
      }
    } catch (err) {
      return actuationError(
        req,
        'write_failed',
        err instanceof Error ? err.message : 'Recipient write failed.',
      );
    }
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: `recipients:${plan.mode}`,
    };
  }

  /**
   * `add-attachment`: attach a base64 file, an https link, or another mail item to the OPEN DRAFT.
   * `planAttachment` validates (and size-caps, never logs) the payload BEFORE any host call; the
   * host-minted attachment id is recorded in the result location when returned. Nothing sends.
   */
  private async applyAttachment(req: ActuationRequest): Promise<ActuationResult> {
    const resolved = this.resolveDraft(req);
    if (resolved.error) return resolved.error;
    const plan = planAttachment(req);
    if (!plan.ok) return actuationError(req, plan.reason, ATTACHMENT_ERROR_MESSAGES[plan.reason]);
    const add = (() => {
      switch (plan.transport) {
        case 'base64':
          return resolved.draft.addFileAttachmentFromBase64Async;
        case 'uri':
          return resolved.draft.addFileAttachmentAsync;
        case 'item':
          return resolved.draft.addItemAttachmentAsync;
      }
    })();
    if (!add) {
      return actuationError(req, 'no_compose', 'Attachment APIs are unavailable on this draft.');
    }
    try {
      const id = await getAsync<string>((cb) =>
        add(plan.value, plan.name, { isInline: plan.isInline }, cb),
      );
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: id ? `draft-attachment:${id}` : 'draft-attachment',
      };
    } catch (err) {
      return actuationError(
        req,
        'write_failed',
        err instanceof Error ? err.message : 'Attachment write failed.',
      );
    }
  }

  /**
   * `set-body`: replace the OPEN DRAFT's body via `Body.setAsync` (HTML default; Text when the
   * payload came from prose). The prior body stays undoable in Outlook; nothing sends.
   */
  private async applyBody(req: ActuationRequest): Promise<ActuationResult> {
    const resolved = this.resolveDraft(req);
    if (resolved.error) return resolved.error;
    const plan = planDraftBody(req);
    if (!plan.data.trim()) {
      return actuationError(req, 'no_body', 'set-body needs params.mail.body or params.text');
    }
    try {
      await getAsync<void>((cb) =>
        resolved.draft.body.setAsync(
          plan.data,
          {
            coercionType:
              plan.coercionType === 'text' ? Office.CoercionType.Text : Office.CoercionType.Html,
          },
          cb,
        ),
      );
    } catch (err) {
      return actuationError(
        req,
        'write_failed',
        err instanceof Error ? err.message : 'Body write failed.',
      );
    }
    return { ok: true, changeId: req.changeId, kind: req.kind, location: 'draft-body' };
  }

  /**
   * `set-subject`: replace the OPEN DRAFT's subject via `Subject.setAsync`. Nothing sends — the
   * user reviews and sends the draft themselves.
   */
  private async applySubject(req: ActuationRequest): Promise<ActuationResult> {
    const resolved = this.resolveDraft(req);
    if (resolved.error) return resolved.error;
    const subject = req.params.mail?.subject?.trim();
    if (!subject) {
      return actuationError(req, 'no_subject', 'set-subject needs params.mail.subject');
    }
    try {
      await getAsync<void>((cb) => resolved.draft.subject.setAsync(subject, cb));
    } catch (err) {
      return actuationError(
        req,
        'write_failed',
        err instanceof Error ? err.message : 'Subject write failed.',
      );
    }
    return { ok: true, changeId: req.changeId, kind: req.kind, location: 'draft-subject' };
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

function outlookItemIdFromRef(ref: ContextRef): string | undefined {
  if (ref.surface !== 'outlook' || ref.kind !== 'mail-item') return undefined;
  if (ref.id && ref.id !== 'outlook:item') return ref.id;
  const item = Office.context.mailbox?.item;
  return item ? readItemId(item) : undefined;
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
 * `Succeeded`, reject on failure. Keeps the bridge body await-flat like `Word.run`. The same
 * `AsyncResult` shape backs the compose write callbacks (`setAsync`/`addAsync`/`add*AttachmentAsync`),
 * so the draft-edit paths reuse this helper.
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

/** Structural slice of a compose-mode `Recipients` field — only what {@link OutlookBridge} drives. */
interface DraftRecipients {
  addAsync(recipients: string[], callback: (result: Office.AsyncResult<void>) => void): void;
  setAsync(recipients: string[], callback: (result: Office.AsyncResult<void>) => void): void;
}

/** Structural slice of a COMPOSE-mode mail item (Office.MessageCompose). `body`/`subject` are the
 * write-capable host objects — read-mode items carry same-named statics without these callbacks,
 * which is exactly what {@link draftView} duck-checks. Attachment methods stay optional so a thin
 * host degrades to a corrective instead of a thrown TypeError. */
interface DraftItem {
  to?: DraftRecipients;
  cc?: DraftRecipients;
  bcc?: DraftRecipients;
  body: {
    setAsync(
      data: string,
      options: { coercionType: unknown },
      callback: (result: Office.AsyncResult<void>) => void,
    ): void;
  };
  subject: {
    setAsync(subject: string, callback: (result: Office.AsyncResult<void>) => void): void;
  };
  addFileAttachmentAsync?(
    uri: string,
    attachmentName: string,
    options: { isInline: boolean },
    callback: (result: Office.AsyncResult<string>) => void,
  ): void;
  addFileAttachmentFromBase64Async?(
    base64File: string,
    attachmentName: string,
    options: { isInline: boolean },
    callback: (result: Office.AsyncResult<string>) => void,
  ): void;
  addItemAttachmentAsync?(
    itemId: string,
    attachmentName: string,
    options: { isInline: boolean },
    callback: (result: Office.AsyncResult<string>) => void,
  ): void;
}

/**
 * Duck-narrow `Office.context.mailbox.item` to a compose-mode draft. Read-mode items share the
 * static shapes but lack the write callbacks, so the presence of `body.setAsync` + `subject.setAsync`
 * is the compose signal — the same function-truthiness guard idiom `revealContext` uses.
 */
function draftView(item: Office.Item | undefined): DraftItem | undefined {
  if (!item) return undefined;
  const candidate = item as unknown as DraftItem;
  const compose =
    typeof candidate.body?.setAsync === 'function' &&
    typeof candidate.subject?.setAsync === 'function';
  return compose ? candidate : undefined;
}

/** Narrow a requested To/Cc/Bcc field to a functional compose `Recipients`; undefined otherwise. */
function draftRecipients(item: DraftItem, field: 'to' | 'cc' | 'bcc'): DraftRecipients | undefined {
  const recipients = item[field];
  return recipients &&
    typeof recipients.addAsync === 'function' &&
    typeof recipients.setAsync === 'function'
    ? recipients
    : undefined;
}

/** Corrective `ActuationResult` for a failed apply — the model can self-fix from code + message. */
function actuationError(req: ActuationRequest, code: string, message: string): ActuationResult {
  return { ok: false, changeId: req.changeId, kind: req.kind, error: { code, message } };
}

/** Why an add-attachment request was rejected before any host call (payloads are never echoed). */
const ATTACHMENT_ERROR_MESSAGES: Record<
  'no_attachment' | 'attachment_too_large' | 'invalid_attachment',
  string
> = {
  no_attachment:
    'add-attachment needs params.attachment base64, https uri, or itemId plus a file name',
  attachment_too_large: 'add-attachment base64 payload exceeds the 3 MB ceiling',
  invalid_attachment: 'add-attachment payload is not valid base64, or its uri is not an https link',
};
