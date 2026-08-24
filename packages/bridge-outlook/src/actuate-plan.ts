import type { ActuationRequest } from '@ge/contracts';

/**
 * Pure translation of an actuation into a host plan — testable without Office.js. A reply/compose
 * draft is built from `params.mail` (explicit to/subject/body) falling back to `params.text` for
 * the body, so an agent that only produced prose still yields a usable draft.
 */
export interface ReplyPlan {
  body: string;
  to?: string[];
  subject?: string;
}

export function planReply(req: ActuationRequest): ReplyPlan {
  const mail = req.params.mail;
  const body = mail?.body ?? req.params.text ?? '';
  return {
    body,
    ...(mail?.to ? { to: mail.to } : {}),
    ...(mail?.subject ? { subject: mail.subject } : {}),
  };
}

/**
 * A brand-new draft (`create-mail`): subject + body from `params.mail` (body falls back to
 * `params.text`). Recipients are passed through only if the agent supplied them — by default a
 * compose draft is left unaddressed for the user to fill, so nothing is auto-sent to anyone.
 */
export interface ComposePlan {
  subject: string;
  body: string;
  to?: string[];
}

export function planCompose(req: ActuationRequest): ComposePlan {
  const mail = req.params.mail;
  return {
    subject: mail?.subject ?? '',
    body: mail?.body ?? req.params.text ?? '',
    ...(mail?.to ? { to: mail.to } : {}),
  };
}

/**
 * `set-recipients`: normalized To/Cc/Bcc targets + the write mode (`set` replaces via setAsync,
 * `add` appends via addAsync). Blank entries are dropped; a plan with zero fields means the
 * request carried no usable address and the bridge corrects with `no_recipients`.
 */
export interface RecipientsPlan {
  mode: 'set' | 'add';
  fields: Array<{ field: 'to' | 'cc' | 'bcc'; addresses: string[] }>;
}

export function planRecipients(req: ActuationRequest): RecipientsPlan {
  const mail = req.params.mail;
  const fields: RecipientsPlan['fields'] = [];
  const collect = (field: 'to' | 'cc' | 'bcc', addresses?: string[]): void => {
    const clean = (addresses ?? []).map((a) => a.trim()).filter((a) => a.length > 0);
    if (clean.length > 0) fields.push({ field, addresses: clean });
  };
  collect('to', mail?.to);
  collect('cc', mail?.cc);
  collect('bcc', mail?.bcc);
  return { mode: mail?.recipientMode ?? 'set', fields };
}

/**
 * `set-body`: payload + coercion for `Body.setAsync`. Explicit `params.mail.body` wins; falls back
 * to `params.html` then `params.text` (mirroring planReply's prose fallback — text infers Text
 * coercion); an explicit `params.mail.coercion` overrides the inference.
 */
export interface DraftBodyPlan {
  data: string;
  coercionType: 'html' | 'text';
}

export function planDraftBody(req: ActuationRequest): DraftBodyPlan {
  const mail = req.params.mail;
  let data = mail?.body ?? '';
  let coercionType: 'html' | 'text' = 'html';
  if (!data && req.params.html !== undefined) data = req.params.html;
  if (!data && req.params.text !== undefined) {
    data = req.params.text;
    coercionType = 'text';
  }
  if (mail?.coercion) coercionType = mail.coercion;
  return { data, coercionType };
}

/**
 * Ceiling for base64 attachment payloads accepted from params (encoded size). Oversized payloads
 * are rejected BEFORE any host call so the bytes never reach Office.js.
 */
export const MAX_ATTACHMENT_BASE64_CHARS = 3 * 1024 * 1024;

/**
 * `add-attachment`: which compose transport carries the payload (`base64` →
 * addFileAttachmentFromBase64Async, `uri` → addFileAttachmentAsync, `item` → addItemAttachmentAsync),
 * under what display name, plus inline-ness. Base64 is whitespace-stripped then charset/padding
 * validated and size-capped; `uri` is screened to https so a local-file or other unsafe scheme can
 * never be handed to the host. The raw payload travels in `value`; it is never logged or echoed.
 */
export type AttachmentPlan =
  | {
      ok: true;
      transport: 'base64' | 'uri' | 'item';
      value: string;
      name: string;
      isInline: boolean;
    }
  | { ok: false; reason: 'no_attachment' | 'attachment_too_large' | 'invalid_attachment' };

const BASE64_RE = /^[A-Za-z0-9+/]+={0,2}$/;

export function planAttachment(req: ActuationRequest): AttachmentPlan {
  const att = req.params.attachment;
  if (!att) return { ok: false, reason: 'no_attachment' };
  const name = att.name?.trim();
  if (typeof att.base64 === 'string') {
    const base64 = att.base64.replace(/\s+/g, '');
    if (!base64 || !name) return { ok: false, reason: 'no_attachment' };
    if (base64.length > MAX_ATTACHMENT_BASE64_CHARS) {
      return { ok: false, reason: 'attachment_too_large' };
    }
    if (base64.length % 4 !== 0 || !BASE64_RE.test(base64)) {
      return { ok: false, reason: 'invalid_attachment' };
    }
    return { ok: true, transport: 'base64', value: base64, name, isInline: att.isInline === true };
  }
  if (typeof att.uri === 'string') {
    const uri = att.uri.trim();
    if (!/^https:\/\//i.test(uri)) return { ok: false, reason: 'invalid_attachment' };
    const derived = decodeURIComponent(uri.split(/[?#]/)[0]!.split('/').pop() ?? '').trim();
    const finalName = (name ?? derived).trim();
    if (!finalName) return { ok: false, reason: 'no_attachment' };
    return {
      ok: true,
      transport: 'uri',
      value: uri,
      name: finalName,
      isInline: att.isInline === true,
    };
  }
  if (typeof att.itemId === 'string' && att.itemId.trim()) {
    if (!name) return { ok: false, reason: 'no_attachment' };
    return { ok: true, transport: 'item', value: att.itemId.trim(), name, isInline: false };
  }
  return { ok: false, reason: 'no_attachment' };
}
