import { officeHostProfile } from './office-host-profile.js';
import { unknownActuationResult } from '@ge/contracts';
import { createBridgeDispatch, negotiateCapabilities } from '@ge/runtime';
import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
  Verification,
} from '@ge/contracts';
import type { AttachmentBytes, DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot, decodeBase64 } from '@ge/content';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import {
  mailAttachmentsToRefs,
  mailAttachmentToContext,
  mailItemToContext,
  mailItemToDocStateBlocks,
  searchMailItem,
  type MailAttachment,
  type MailAttachmentContent,
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

// No quotes: the snapshot renderer escapes them, which garbles a command example.
// mail/compose open a separate window the add-in cannot edit or address (recipients are never
// auto-filled): live 2026-10-09 the model followed them with set-recipients/set-body and repeats.
// At most 240 characters (the snapshot renderer clips a summary there); the sender comes first so a
// request to someone else reads as a new email, not a reply (live 2026-10-09).
const receivedNote = (from: string | undefined): string =>
  `received message${plainAddress(from) ? ` from ${plainAddress(from)}` : ''}: reply = mail (a file: attach=); email to anyone else = compose. ` +
  'Each opens a window the user reviews (no later edits): one completes it. set-* and add-attachment: drafts only';

/** The sender only when it is a plain address: a quoted local part could carry `)` or `;`. */
function plainAddress(from: string | undefined): string | undefined {
  return from && /^[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9.-]{1,253}$/.test(from) && from.length <= 40
    ? from
    : undefined;
}
const DRAFT_NOTE =
  'a draft being written: set-subject, set-body, set-recipients and add-attachment edit it';
const DRAFT_READ_TIMEOUT_MS = 2000;
/**
 * A draft write (`setAsync`/`addAsync`/`add*AttachmentAsync`) that never calls back would otherwise
 * hang the task forever. Past this budget the write is reported as an UNKNOWN outcome — the host may
 * still have applied it — never as a success and never as a clean failure.
 */
const DRAFT_WRITE_TIMEOUT_MS = 10_000;

/** `isSetSupported` through optional access: a thin or non-mailbox host may lack the probe. */
function mailboxSetSupported(version: string): boolean {
  try {
    const probe = (
      Office.context as { requirements?: { isSetSupported?(name: string, v: string): boolean } }
    ).requirements?.isSetSupported;
    return probe ? probe.call(Office.context.requirements, 'Mailbox', version) === true : false;
  } catch {
    return false;
  }
}

export class OutlookBridge implements DocBridge {
  private static readonly dispatcher = createBridgeDispatch<OutlookBridge>(
    'outlook',
    {
      'reply-mail': (host, request) => host.applyReply(request),
      'create-mail': (host, request) => host.applyCompose(request),
      'set-recipients': (host, request) => host.applyRecipients(request),
      'add-attachment': (host, request) => host.applyAttachment(request),
      'set-body': (host, request) => host.applyBody(request),
      'set-subject': (host, request) => host.applySubject(request),
    },
    {
      provenance: 'unsupported',
      capabilities: (host) => host.getCapabilities(),
      documentKey: (host) => host.documentKey(),
      unavailable: (host, request) =>
        ['set-body', 'set-subject', 'set-recipients', 'add-attachment'].includes(request.kind)
          ? host.resolveDraft(request).error
          : undefined,
    },
  );
  static readonly handledActuations = OutlookBridge.dispatcher.handledActuations;

  readonly surface = 'outlook' as const;

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  getHostProfile(): import('@ge/contracts').HostCapabilityProfile | undefined {
    return officeHostProfile(this.surface);
  }

  getCapabilities(): CapabilityManifest {
    return negotiateCapabilities(OUTLOOK_CAPABILITIES, this.getHostProfile());
  }

  async listContext(): Promise<ContextRef[]> {
    const item = Office.context.mailbox?.item;
    if (!item) return [];
    const subject = await readCurrentSubject(item);
    if (Office.context.mailbox?.item !== item)
      throw new Error('The active mail item changed during capture.');
    return [
      {
        id: item.itemId ?? 'outlook:item',
        kind: 'mail-item',
        surface: 'outlook',
        title: subject ? `Email: ${subject}` : 'Current email',
        ...(subject ? { preview: subject.slice(0, 120) } : {}),
        live: true,
      },
      // Read-mode attachments (Mailbox 1.1 `item.attachments`): listed by name, content not read.
      ...mailAttachmentsToRefs(item.itemId, readAttachmentDetails(item)),
    ];
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.surface === this.surface && ref.kind === 'attachment')
      return this.resolveAttachment(ref);
    if (ref.surface !== this.surface || ref.kind !== 'mail-item')
      throw new Error('Only the active Outlook mail item can be captured.');
    const current = Office.context.mailbox?.item;
    if (current && ref.id !== 'outlook:item' && ref.id !== readItemId(current))
      throw new Error(
        'The requested mail item is no longer active. Attach the current item again.',
      );
    const mail = await this.readMailItem();
    if (mail && ref.id !== 'outlook:item' && ref.id !== mail.id)
      throw new Error(
        'The requested mail item is no longer active. Attach the current item again.',
      );
    return mail ? mailItemToContext(mail) : [];
  }

  /**
   * One attachment of the ACTIVE item → context through the shared attachment pipeline. Text-like
   * files, forwarded messages (`.eml`) and invites are read with `getAttachmentContentAsync`
   * (Mailbox 1.8) and inlined as data; PDF/Office files become one uploadable `file` part the
   * runtime materializes on an explicit attach; without 1.8 the attachment grounds by name.
   */
  private async resolveAttachment(ref: ContextRef): Promise<ResolvedContext[]> {
    const { item, attachment } = this.activeAttachment(ref);
    if (!attachment) return [];
    const content = await this.readAttachmentContent(item, attachment.id);
    if (Office.context.mailbox?.item !== item)
      throw new Error('The active mail item changed during capture.');
    return mailAttachmentToContext(readItemId(item), attachment, content);
  }

  /**
   * Raw bytes of a file attachment for the runtime's session-context upload (explicit user attach
   * only; never from a model command). Requires Mailbox 1.8; the bytes are returned, not kept.
   */
  async readAttachment(ref: ContextRef): Promise<AttachmentBytes> {
    const { item, attachment } = this.activeAttachment(ref);
    if (!attachment) throw new Error('This attachment is not on the active mail item.');
    const content = await this.readAttachmentContent(item, attachment.id);
    if (!content || content.format !== 'base64')
      throw new Error('This host cannot read the attachment content.');
    const bytes = decodeBase64(content.content);
    return {
      fileName: attachment.name,
      mimeType: attachment.contentType ?? 'application/octet-stream',
      bytes,
    };
  }

  /** The active item plus the attachment a ref points at, matched by host pointer and item. */
  private activeAttachment(ref: ContextRef): {
    item: Office.Item;
    attachment: MailAttachment | undefined;
  } {
    const item = Office.context.mailbox?.item;
    if (!item) throw new Error('No active mail item.');
    const pointer = ref.hostRef?.type === 'outlook.attachment' ? ref.hostRef : undefined;
    if (!pointer) return { item, attachment: undefined };
    const itemId = readItemId(item) ?? 'item';
    if (pointer.itemId !== itemId)
      throw new Error(
        'The requested mail item is no longer active. Attach the current item again.',
      );
    const attachment = readAttachmentDetails(item).find((a) => a.id === pointer.attachmentId);
    return { item, attachment };
  }

  private async readAttachmentContent(
    item: Office.Item,
    attachmentId: string,
  ): Promise<MailAttachmentContent | undefined> {
    const reader = (item as { getAttachmentContentAsync?: unknown }).getAttachmentContentAsync;
    if (
      typeof reader !== 'function' ||
      !(Office.context.requirements?.isSetSupported?.('Mailbox', '1.8') ?? false)
    )
      return undefined;
    try {
      const result = await getAsync<{ format: string; content: string }>(
        (cb) =>
          (reader as (id: string, cb: (r: Office.AsyncResult<unknown>) => void) => void).call(
            item,
            attachmentId,
            cb as (r: Office.AsyncResult<unknown>) => void,
          ),
        ATTACHMENT_READ_TIMEOUT_MS,
      );
      return result && typeof result.content === 'string'
        ? { format: result.format, content: result.content }
        : undefined;
    } catch {
      // A failed or timed-out read grounds the attachment by name; it never breaks the attach.
      return undefined;
    }
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
  documentKey(): string | undefined {
    try {
      return Office.context?.mailbox?.item?.itemId || undefined;
    } catch {
      return undefined;
    }
  }

  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    const mail = await this.readMailItem();
    if (!mail) return undefined;
    const blocks = mailItemToDocStateBlocks(mail);
    if (blocks.length === 0) return undefined;
    this.docStateVersion += 1;
    const snapshot = buildDocStateSnapshot({
      surface: 'outlook',
      version: this.docStateVersion,
      ...(mail.subject?.trim() ? { title: mail.subject } : {}),
      blocks,
    });
    // Say up front which writes this item takes. Without it the model learned only from a failed
    // write: it edited a received message's body/recipients, and used compose (a NEW message) for a
    // reply (live scenarios 2026-10-09).
    const item = Office.context.mailbox?.item;
    const entry = snapshot.inventory.find((e) => e.id === 'mail:item') ?? snapshot.inventory[0];
    if (item && entry) entry.summary = draftView(item) ? DRAFT_NOTE : receivedNote(mail.from);
    return snapshot;
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
    const [body, subject] = await Promise.all([
      getAsync<string>(
        (cb) => item.body.getAsync(Office.CoercionType.Html, cb),
        DRAFT_READ_TIMEOUT_MS,
      ),
      readCurrentSubject(item),
    ]);
    if (Office.context.mailbox?.item !== item)
      throw new Error('The active mail item changed during capture.');
    return {
      ...(item.itemId ? { id: item.itemId } : {}),
      ...(subject ? { subject } : {}),
      ...(readFrom(item) ? { from: readFrom(item) } : {}),
      body,
      bodyType: 'html',
    };
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    return OutlookBridge.dispatcher.dispatch(this, req);
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
    // `ItemChanged` arrived with Mailbox 1.5. A host below that (or one that cannot answer the
    // requirement probe at all) gets a no-op subscription instead of a thrown registration.
    if (!mailboxSetSupported('1.5')) return () => {};

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
    // A file opened with the reply (`mail … attach=`): a received message cannot take
    // add-attachment, but its reply form can carry a base64 file (Mailbox 1.15).
    let attachments: Array<{ type: string; name: string; base64file: string }> | undefined;
    if (req.params.attachment) {
      // Same checks as add-attachment: a base64 file with a name, under the 3 MB ceiling.
      const file = planAttachment(req);
      if (!file.ok) return actuationError(req, file.reason, ATTACHMENT_ERROR_MESSAGES[file.reason]);
      if (file.transport !== 'base64')
        return actuationError(req, 'invalid_attachment', 'A reply carries only a built file.');
      attachments = [{ type: 'base64', name: file.name, base64file: file.value }];
    }
    if (attachments && !Office.context.requirements?.isSetSupported?.('Mailbox', '1.15')) {
      return actuationError(
        req,
        'unsupported',
        'This Outlook version cannot open a reply with a file attached (needs Mailbox 1.15). ' +
          'Nothing was opened. Reply without attach= and tell the user to attach the file themselves.',
      );
    }
    // Open a reviewable reply form rather than sending — the user confirms before it leaves.
    item.displayReplyForm({
      htmlBody: plan.body,
      ...(attachments ? { attachments: attachments as Office.ReplyFormAttachment[] } : {}),
    });
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: attachments ? 'reply-form+attachment' : 'reply-form',
      completion: 'review-opened',
    };
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
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: 'new-message-form',
      completion: 'review-opened',
    };
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
            // Say what to do instead: without it the model tried every draft verb in turn, then
            // `compose` (a NEW message) when a reply was wanted. But a bare "use mail" hint made it
            // answer "change the subject" of a draft in another window by opening an unrelated
            // reply (live 2026-10-14), so the hint is conditional on what the user asked for.
            'The open item is a received message, not a draft being written, so set-subject, ' +
              'set-body, set-recipients and add-attachment cannot run here. A draft opened in its ' +
              'own window cannot be edited from this pane: do not open a reply or a new message ' +
              'instead; emit done and tell the user to make the change in that draft window, or ' +
              'to open this add-in from it. Only if the user asked to reply: mail "reply text" (with a ' +
              'file: mail "reply text" attach="name.docx" text="file text"). ' +
              'Only if they asked for a new email: compose "Subject" "body".',
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
        if (Office.context.mailbox?.item !== (resolved.draft as unknown as Office.Item))
          throw new Error('The active draft changed during the recipient update.');
        await getAsync<void>(
          (cb) =>
            plan.mode === 'add'
              ? recipients.addAsync(addresses, cb)
              : recipients.setAsync(addresses, cb),
          DRAFT_WRITE_TIMEOUT_MS,
        );
      }
    } catch (err) {
      return unknownActuationResult(
        req,
        `Outlook did not confirm the dispatched change. Inspect the draft before trying again. ${err instanceof Error ? err.message : ''}`.trim(),
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
      // Call ON the draft: Office.js methods need their item as `this`, and a detached call fails.
      const id = await getAsync<string>(
        (cb) => add.call(resolved.draft, plan.value, plan.name, { isInline: plan.isInline }, cb),
        DRAFT_WRITE_TIMEOUT_MS,
      );
      return {
        ok: true,
        changeId: req.changeId,
        kind: req.kind,
        location: id ? `draft-attachment:${id}` : 'draft-attachment',
      };
    } catch (err) {
      return unknownActuationResult(
        req,
        `Outlook did not confirm the dispatched change. Inspect the draft before trying again. ${err instanceof Error ? err.message : ''}`.trim(),
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
      await getAsync<void>(
        (cb) =>
          resolved.draft.body.setAsync(
            plan.data,
            {
              coercionType:
                plan.coercionType === 'text' ? Office.CoercionType.Text : Office.CoercionType.Html,
            },
            cb,
          ),
        DRAFT_WRITE_TIMEOUT_MS,
      );
    } catch (err) {
      return unknownActuationResult(
        req,
        `Outlook did not confirm the dispatched change. Inspect the draft before trying again. ${err instanceof Error ? err.message : ''}`.trim(),
      );
    }
    const verification = await verifyDraftField(resolved.draft, async () => {
      if (!resolved.draft.body.getAsync) return undefined;
      const text = await getAsync<string>(
        (cb) =>
          resolved.draft.body.getAsync!(
            plan.coercionType === 'text' ? Office.CoercionType.Text : Office.CoercionType.Html,
            cb,
          ),
        DRAFT_READ_TIMEOUT_MS,
      );
      // HTML hosts may normalize wrappers/markup: an exact match proves the write; a differing
      // serialized HTML form does not prove a substantive mismatch.
      return text === plan.data
        ? 'verified'
        : plan.coercionType === 'text'
          ? 'mismatch'
          : 'unknown';
    });
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: 'draft-body',
      verification,
    };
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
      await getAsync<void>(
        (cb) => resolved.draft.subject.setAsync(subject, cb),
        DRAFT_WRITE_TIMEOUT_MS,
      );
    } catch (err) {
      return unknownActuationResult(
        req,
        `Outlook did not confirm the dispatched change. Inspect the draft before trying again. ${err instanceof Error ? err.message : ''}`.trim(),
      );
    }
    const verification = await verifyDraftField(resolved.draft, async () => {
      if (!resolved.draft.subject.getAsync) return undefined;
      return (await getAsync<string>(
        (cb) => resolved.draft.subject.getAsync!(cb),
        DRAFT_READ_TIMEOUT_MS,
      )) === subject
        ? 'verified'
        : 'mismatch';
    });
    return {
      ok: true,
      changeId: req.changeId,
      kind: req.kind,
      location: 'draft-subject',
      verification,
    };
  }
}

/**
 * Classify the newly active mail item into a HostEvent. Compose APIs identify drafts even
 * after saving mints an `itemId`; received items carry `mail-received`. Returns
 * undefined when no item is active (nothing to emit). Reads only the opaque `itemId`, never the
 * subject/body (untrusted content stays out of the event).
 */
function activeItemEvent(item: Office.Item | undefined): HostEvent | undefined {
  if (!item) return undefined;
  const id = readItemId(item);
  if (draftView(item)) return composeEvent(id);
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

async function readCurrentSubject(item: Office.Item): Promise<string | undefined> {
  const draft = draftView(item);
  if (draft?.subject.getAsync)
    return getAsync<string>((cb) => draft.subject.getAsync!(cb), DRAFT_READ_TIMEOUT_MS);
  return readSubject(item);
}

async function verifyDraftField(
  draft: DraftItem,
  read: () => Promise<Verification['status'] | undefined>,
): Promise<Verification> {
  try {
    const status = await read();
    if (Office.context.mailbox?.item !== (draft as unknown as Office.Item))
      return {
        status: 'unknown',
        message: 'The active item changed after dispatch. Inspect the draft.',
      };
    return { status: status ?? 'unknown' };
  } catch {
    return { status: 'unknown', message: 'The draft write returned but readback was unavailable.' };
  }
}

/** Bound on one `getAttachmentContentAsync` round trip (large files on slow hosts). */
const ATTACHMENT_READ_TIMEOUT_MS = 10_000;

/** `item.attachments` as plain data (read mode); an item without the collection has none. */
function readAttachmentDetails(item: Office.Item): MailAttachment[] {
  const raw: unknown = (item as { attachments?: unknown }).attachments;
  if (!Array.isArray(raw)) return [];
  return raw.flatMap((a: unknown): MailAttachment[] => {
    if (!a || typeof a !== 'object') return [];
    const d = a as Record<string, unknown>;
    if (typeof d.id !== 'string' || typeof d.name !== 'string') return [];
    return [
      {
        id: d.id,
        name: d.name,
        ...(typeof d.attachmentType === 'string' ? { attachmentType: d.attachmentType } : {}),
        ...(typeof d.contentType === 'string' ? { contentType: d.contentType } : {}),
        ...(typeof d.size === 'number' ? { size: d.size } : {}),
        ...(typeof d.isInline === 'boolean' ? { isInline: d.isInline } : {}),
      },
    ];
  });
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
  timeoutMs?: number,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer =
      timeoutMs === undefined
        ? undefined
        : setTimeout(() => reject(new Error('Outlook host call timed out.')), timeoutMs);
    try {
      call((result) => {
        if (timer !== undefined) clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.error?.message ?? 'Outlook getAsync failed'));
        }
      });
    } catch (error) {
      if (timer !== undefined) clearTimeout(timer);
      reject(error);
    }
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
    getAsync?(coercionType: unknown, callback: (result: Office.AsyncResult<string>) => void): void;
    setAsync(
      data: string,
      options: { coercionType: unknown },
      callback: (result: Office.AsyncResult<void>) => void,
    ): void;
  };
  subject: {
    getAsync?(callback: (result: Office.AsyncResult<string>) => void): void;
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
    'add-attachment needs a file name plus one of: text="plain text", base64=<base64>, or uri="https://…" — e.g. /add-attachment name="hello.txt" text="Hello"',
  attachment_too_large: 'add-attachment base64 payload exceeds the 3 MB ceiling',
  invalid_attachment: 'add-attachment payload is not valid base64, or its uri is not an https link',
};

/** Actual dispatch keys; conformance checks these against the advertised capabilities. */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = OutlookBridge.handledActuations;
