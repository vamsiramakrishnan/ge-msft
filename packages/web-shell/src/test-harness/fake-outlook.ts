/**
 * In-memory **Outlook host simulator**. Models the slice of `Office.context.mailbox` (+ the active
 * `Office.context.mailbox.item`) the real {@link "@ge/bridge-outlook"!OutlookBridge} and the on-send
 * glue ({@link "@ge/bridge-outlook"!createMessageSendHandler}) drive — the bridge's only un-faked
 * seam — so the REAL bridge + REAL on-send handler run unchanged against a seeded mail item.
 *
 * Unlike Word/Excel (`Word.run` / `Excel.run` over a native object model), Outlook is a single
 * `Office.context.mailbox.item` whose reads are **callback-based** (`getAsync`), and whose writes
 * open **reviewable host forms** (`displayReplyForm` / `displayNewMessageForm`) rather than mutating
 * silently. The on-send gate is the consequential surface: a Smart Alerts `OnMessageSend` event
 * whose `completed({ allowEvent, errorMessage })` decides whether the mail leaves.
 *
 * Enumerated host calls modelled (the fidelity boundary for Outlook):
 *   - `Office.context.mailbox` — the mailbox; `undefined` for a non-mailbox host (defensive paths).
 *   - `Office.context.mailbox.item` — the active item: `itemId` / `subject` / `from.emailAddress`.
 *   - `item.body.getAsync(Office.CoercionType.Html, cb)` → the HTML body (the callback read path).
 *   - `item.displayReplyForm({ htmlBody })` (the reviewable reply WRITE) — recorded as a draft form.
 *   - `mailbox.displayNewMessageForm({ subject, htmlBody, toRecipients })` (the compose WRITE).
 *   - `mailbox.addHandlerAsync/removeHandlerAsync(Office.EventType.ItemChanged, h)` (watch()).
 *   - `Office.CoercionType.Html`, `Office.AsyncResultStatus.Succeeded`, `Office.EventType.ItemChanged`.
 *
 * Fidelity notes / boundary:
 *   - `getAsync` resolves on the next microtask with `{ status: 'succeeded', value }` so the bridge's
 *     `promisify` wrapper exercises its real success path (or `failed` when the seed forces a read
 *     error). The host content (subject/from/body) is UNTRUSTED — it flows through the bridge as data.
 *   - The reply/compose WRITES never "send": they record the opened form in the seed, mirroring the
 *     host's reviewable-form contract. Nothing leaves the client without the user, and — crucially —
 *     the on-send gate (separate from these forms) decides every actual send.
 */

import { installGlobal } from './globals.js';

/** The sender shape Outlook exposes in read mode (`item.from.emailAddress`). */
export interface MailFromSeed {
  emailAddress: string;
  displayName?: string;
}

/** A recorded reviewable reply form the bridge opened via `displayReplyForm`. */
export interface ReplyFormSeed {
  htmlBody: string;
}

/** A recorded compose draft the bridge opened via `displayNewMessageForm`. */
export interface NewMessageFormSeed {
  subject: string;
  htmlBody: string;
  toRecipients?: string[];
}

/** The Outlook mail-item seed: the active item's header fields + body, plus recorded writes. */
export interface OutlookSeed {
  /** The saved item id (read mode). Undefined/empty for an unsaved compose draft. */
  itemId?: string;
  subject?: string;
  from?: MailFromSeed;
  /** The HTML body returned by `body.getAsync(Html, …)`. */
  body: string;
  /** When true, `body.getAsync` resolves `failed` so the bridge's read-error path is exercised. */
  bodyReadFails?: boolean;
  /** Recorded reviewable reply forms opened via `displayReplyForm`. */
  replyForms: ReplyFormSeed[];
  /** Recorded compose drafts opened via `displayNewMessageForm`. */
  newMessageForms: NewMessageFormSeed[];
}

/* ─────────────────────────── fake object model ─────────────────────────── */

const COERCION_TYPE = { Html: 'html', Text: 'text' } as const;
const ASYNC_RESULT_STATUS = { Succeeded: 'succeeded', Failed: 'failed' } as const;
const EVENT_TYPE = { ItemChanged: 'olkItemChanged' } as const;

interface FakeAsyncResult<T> {
  status: string;
  value: T;
  error?: { message: string };
}

class FakeBody {
  constructor(private readonly seed: OutlookSeed) {}
  getAsync(_coercionType: string, callback: (result: FakeAsyncResult<string>) => void): void {
    // Mirror the host: resolve on the next microtask so the bridge's promisify await is real.
    queueMicrotask(() => {
      if (this.seed.bodyReadFails) {
        callback({
          status: ASYNC_RESULT_STATUS.Failed,
          value: '',
          error: { message: 'simulated body read failure' },
        });
      } else {
        callback({ status: ASYNC_RESULT_STATUS.Succeeded, value: this.seed.body });
      }
    });
  }
}

/** The active mail item: read-mode header fields + the callback body + the reviewable reply WRITE. */
class FakeMailItem {
  readonly body: FakeBody;
  constructor(private readonly seed: OutlookSeed) {
    this.body = new FakeBody(seed);
  }
  get itemId(): string | undefined {
    return this.seed.itemId && this.seed.itemId.length > 0 ? this.seed.itemId : undefined;
  }
  get subject(): string | undefined {
    return this.seed.subject;
  }
  get from(): MailFromSeed | undefined {
    return this.seed.from;
  }
  /** Reviewable reply WRITE — record the opened form rather than sending. */
  displayReplyForm(options: { htmlBody: string }): void {
    this.seed.replyForms.push({ htmlBody: options.htmlBody });
  }
}

type ItemChangedHandler = () => void;

class FakeMailbox {
  readonly item: FakeMailItem;
  private readonly handlers = new Map<string, ItemChangedHandler[]>();
  constructor(private readonly seed: OutlookSeed) {
    this.item = new FakeMailItem(seed);
  }
  /** Compose WRITE — open a brand-new reviewable draft; record it (never auto-send). */
  displayNewMessageForm(options: {
    subject: string;
    htmlBody: string;
    toRecipients?: string[];
  }): void {
    this.seed.newMessageForms.push({
      subject: options.subject,
      htmlBody: options.htmlBody,
      ...(options.toRecipients ? { toRecipients: options.toRecipients } : {}),
    });
  }
  addHandlerAsync(
    eventType: string,
    handler: ItemChangedHandler,
    callback?: (result: { status: string }) => void,
  ): void {
    const list = this.handlers.get(eventType) ?? [];
    list.push(handler);
    this.handlers.set(eventType, list);
    callback?.({ status: ASYNC_RESULT_STATUS.Succeeded });
  }
  removeHandlerAsync(
    eventType: string,
    handler?: ItemChangedHandler,
    callback?: (result: { status: string }) => void,
  ): void {
    const list = this.handlers.get(eventType) ?? [];
    this.handlers.set(eventType, handler ? list.filter((h) => h !== handler) : []);
    callback?.({ status: ASYNC_RESULT_STATUS.Succeeded });
  }
  /** Fire every handler registered for `eventType` (drives `watch()` in a test). */
  fire(eventType: string): void {
    for (const h of this.handlers.get(eventType) ?? []) h();
  }
  count(eventType: string): number {
    return (this.handlers.get(eventType) ?? []).length;
  }
}

/** The fake `Office` namespace object installed onto `globalThis.Office` for the Outlook surface. */
interface FakeOutlookOffice {
  context: { mailbox?: FakeMailbox };
  CoercionType: typeof COERCION_TYPE;
  AsyncResultStatus: typeof ASYNC_RESULT_STATUS;
  EventType: typeof EVENT_TYPE;
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** A read-back view of the Outlook host after a run. */
export interface OutlookSnapshot {
  itemId?: string;
  subject?: string;
  body: string;
  replyForms: ReadonlyArray<ReplyFormSeed>;
  newMessageForms: ReadonlyArray<NewMessageFormSeed>;
}

/** The installed Outlook simulator. */
export interface OutlookSimulator {
  readonly seed: OutlookSeed;
  /** The live mailbox — a test can `fire(ItemChanged)` to drive `watch()`. */
  readonly mailbox: FakeMailbox;
  /** The Office `EventType` constants (so a test can fire the right event name). */
  readonly eventType: typeof EVENT_TYPE;
  snapshot(): OutlookSnapshot;
  restore(): void;
}

/**
 * Install an in-memory Outlook host: writes `globalThis.Office` (with a `context.mailbox` + an
 * active `item`) so the REAL {@link "@ge/bridge-outlook"!OutlookBridge} and on-send glue run against
 * `seed`. Pass `{ mailbox: false }` to model a non-mailbox host (`Office.context.mailbox` undefined)
 * so the bridge's defensive no-item paths can be exercised.
 */
export function installFakeOutlook(
  seed: OutlookSeed = defaultOutlookSeed(),
  opts: { mailbox?: boolean } = {},
): OutlookSimulator {
  const mailbox = new FakeMailbox(seed);
  const office: FakeOutlookOffice = {
    context: { ...(opts.mailbox === false ? {} : { mailbox }) },
    CoercionType: COERCION_TYPE,
    AsyncResultStatus: ASYNC_RESULT_STATUS,
    EventType: EVENT_TYPE,
  };

  const restore = installGlobal('Office', office);

  return {
    seed,
    mailbox,
    eventType: EVENT_TYPE,
    snapshot: () => ({
      ...(seed.itemId ? { itemId: seed.itemId } : {}),
      ...(seed.subject ? { subject: seed.subject } : {}),
      body: seed.body,
      replyForms: seed.replyForms.map((f) => ({ ...f })),
      newMessageForms: seed.newMessageForms.map((f) => ({
        ...f,
        ...(f.toRecipients ? { toRecipients: [...f.toRecipients] } : {}),
      })),
    }),
    restore,
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** Build an {@link OutlookSeed}, defaulting empty write logs. */
export function outlookSeed(init: {
  itemId?: string;
  subject?: string;
  from?: MailFromSeed;
  body: string;
  bodyReadFails?: boolean;
}): OutlookSeed {
  return {
    ...(init.itemId ? { itemId: init.itemId } : {}),
    ...(init.subject ? { subject: init.subject } : {}),
    ...(init.from ? { from: init.from } : {}),
    body: init.body,
    ...(init.bodyReadFails ? { bodyReadFails: init.bodyReadFails } : {}),
    replyForms: [],
    newMessageForms: [],
  };
}

/** A realistic inbound mail fixture: a vendor asking for a contract concession (something to reply to). */
export function defaultOutlookSeed(): OutlookSeed {
  return outlookSeed({
    itemId: 'AAMkAGI-msa-thread-001',
    subject: 'Re: Master Services Agreement — SLA and liability cap',
    from: { emailAddress: 'procurement@northwind.example', displayName: 'Northwind Procurement' },
    body:
      '<p>Hi team,</p>' +
      '<p>We can accept the 99.5% SLA, but we need the liability cap raised to 12 months of fees ' +
      'before we can sign. Can you confirm by Friday?</p>' +
      '<p>Thanks,<br/>Procurement</p>',
  });
}
