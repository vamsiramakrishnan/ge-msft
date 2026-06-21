import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import { OUTLOOK_CAPABILITIES } from './capabilities.js';
import { mailItemToContext, type MailItem } from './capture.js';
import { planReply } from './actuate-plan.js';

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
