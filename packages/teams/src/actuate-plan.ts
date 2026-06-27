import type { ActuationRequest } from '@ge/contracts';

/**
 * Pure translation of a Teams actuation into a host post plan — testable without TeamsJS. A
 * reviewable chat post is built from `params.text` (the synthesized prose) with an optional
 * Adaptive Card payload threaded through `params.html` when an agent produced rich card markup.
 * Mirrors Outlook's `planReply`: an agent that only produced prose still yields a usable post.
 *
 * The result is staged for review (never auto-sent) by the bridge.
 */
export interface PostMessagePlan {
  text: string;
  card?: unknown;
}

export function planPostMessage(req: ActuationRequest): PostMessagePlan {
  const text = req.params.text ?? '';
  const card = req.params.html;
  return {
    text,
    ...(card !== undefined ? { card } : {}),
  };
}
