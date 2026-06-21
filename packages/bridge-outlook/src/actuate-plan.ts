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
