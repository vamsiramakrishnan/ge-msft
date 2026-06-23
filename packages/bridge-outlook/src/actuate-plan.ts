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
