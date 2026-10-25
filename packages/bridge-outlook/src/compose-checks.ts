import { contentHash } from '@ge/contracts';
import type { SendDecision } from './events.js';

/** Complete, bounded facts read from the current compose item. No attachment contents or tokens. */
export interface ComposeFacts {
  /** Process-local identity also pins unsaved drafts with identical content. Never a bearer capability. */
  instanceId: string;
  itemId?: string;
  subject: string;
  body: string;
  to: string[];
  cc: string[];
  bcc: string[];
  attachments: Array<{ id: string; name: string; isInline: boolean }>;
}

export interface ComposeSendPolicy {
  requireSubject?: boolean;
  requireRecipient?: boolean;
  /** Exact filenames expected by this deployment/workflow, not instructions extracted from mail. */
  requiredAttachments?: readonly string[];
  /** Optional exact domain allowlist. Subdomains do not inherit permission. */
  allowedRecipientDomains?: readonly string[];
}

export const COMPOSE_FACT_LIMITS = Object.freeze({
  bodyBytes: 64 * 1024,
  subjectBytes: 1024,
  recipients: 100,
  attachments: 50,
  timeoutMs: 2000,
});
const composeInstances = new WeakMap<object, string>();
let nextComposeInstance = 0;

/** Deterministic checks only. This does not assert model grounding, delivery, or tenant DLP. */
export function checkComposeFacts(facts: ComposeFacts, policy: ComposeSendPolicy): SendDecision {
  if (policy.requireSubject && !facts.subject.trim())
    return { allowEvent: false, message: 'Add a subject before sending this message.' };
  const recipients = [...facts.to, ...facts.cc, ...facts.bcc];
  if (policy.requireRecipient && recipients.length === 0)
    return { allowEvent: false, message: 'Add a recipient before sending this message.' };
  if (policy.allowedRecipientDomains) {
    const domains = new Set(policy.allowedRecipientDomains.map((domain) => domain.toLowerCase()));
    if (
      recipients.some((address) => {
        const match = /^[^\s@]+@([^\s@]+)$/.exec(address);
        return !match || !domains.has(match[1]!.toLowerCase());
      })
    )
      return {
        allowEvent: false,
        message: 'A recipient is outside the permitted domains. Review To, Cc and Bcc.',
      };
  }
  if (
    policy.requiredAttachments?.some(
      (name) =>
        !facts.attachments.some((attachment) => !attachment.isInline && attachment.name === name),
    )
  )
    return {
      allowEvent: false,
      message: 'A required file is missing. Review the attachments before sending.',
    };
  return { allowEvent: true };
}

export async function composeFactsFingerprint(facts: ComposeFacts): Promise<string> {
  return contentHash(facts);
}

/** Read actual compose APIs. Every required field must be available; never infer an empty value. */
export async function captureComposeFacts(): Promise<ComposeFacts> {
  const item = Office.context.mailbox?.item as
    | (Office.MessageCompose & { itemId?: string })
    | undefined;
  if (
    !item ||
    typeof item.subject?.getAsync !== 'function' ||
    typeof item.body?.getAsync !== 'function' ||
    typeof item.to?.getAsync !== 'function' ||
    typeof item.cc?.getAsync !== 'function' ||
    typeof item.bcc?.getAsync !== 'function' ||
    typeof item.getAttachmentsAsync !== 'function' ||
    !Office.context.requirements?.isSetSupported('Mailbox', '1.8')
  )
    throw new Error('Complete compose checks require Mailbox 1.8 compose APIs.');
  const id = item.itemId || undefined;
  let instanceId = composeInstances.get(item);
  if (!instanceId) {
    instanceId = `compose_${++nextComposeInstance}`;
    composeInstances.set(item, instanceId);
  }
  const [subject, body, to, cc, bcc, attachments] = await Promise.all([
    read<string>((cb) => item.subject.getAsync(cb)),
    read<string>((cb) => item.body.getAsync(Office.CoercionType.Text, cb)),
    read<Office.EmailAddressDetails[]>((cb) => item.to.getAsync(cb)),
    read<Office.EmailAddressDetails[]>((cb) => item.cc.getAsync(cb)),
    read<Office.EmailAddressDetails[]>((cb) => item.bcc.getAsync(cb)),
    read<Office.AttachmentDetailsCompose[]>((cb) => item.getAttachmentsAsync(cb)),
  ]);
  if (Office.context.mailbox?.item !== item || (item.itemId || undefined) !== id)
    throw new Error('The active draft changed during send checks.');
  const bytes = (value: string) => new TextEncoder().encode(value).byteLength;
  if (
    typeof subject !== 'string' ||
    typeof body !== 'string' ||
    bytes(subject) > COMPOSE_FACT_LIMITS.subjectBytes ||
    bytes(body) > COMPOSE_FACT_LIMITS.bodyBytes ||
    ![to, cc, bcc, attachments].every(Array.isArray) ||
    to.length + cc.length + bcc.length > COMPOSE_FACT_LIMITS.recipients ||
    attachments.length > COMPOSE_FACT_LIMITS.attachments
  )
    throw new Error('The draft exceeds the complete send-check budget.');
  const addresses = (values: Office.EmailAddressDetails[]) =>
    values.map((value) => {
      if (
        typeof value.emailAddress !== 'string' ||
        !value.emailAddress.trim() ||
        bytes(value.emailAddress) > 512
      )
        throw new Error('A recipient could not be resolved completely.');
      return value.emailAddress.trim();
    });
  return {
    instanceId,
    ...(id ? { itemId: id } : {}),
    subject,
    body,
    to: addresses(to),
    cc: addresses(cc),
    bcc: addresses(bcc),
    attachments: attachments.map((attachment) => {
      if (
        typeof attachment.id !== 'string' ||
        typeof attachment.name !== 'string' ||
        typeof attachment.isInline !== 'boolean' ||
        bytes(attachment.id) > 1024 ||
        bytes(attachment.name) > 1024
      )
        throw new Error('An attachment could not be inspected completely.');
      return { id: attachment.id, name: attachment.name, isInline: attachment.isInline };
    }),
  };
}

function read<T>(invoke: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error('Compose capture timed out.')),
      COMPOSE_FACT_LIMITS.timeoutMs,
    );
    try {
      invoke((result) => {
        clearTimeout(timer);
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error('A required compose field could not be read.'));
      });
    } catch (error) {
      clearTimeout(timer);
      reject(error);
    }
  });
}
