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

/**
 * The command scanner turns a quoted `\n` into a line break, which is invalid inside a JSON string,
 * so a card written as `html="{\"text\":\"a\nb\"}"` would no longer parse. Re-escape line breaks
 * inside the card's string literals only; whitespace between JSON tokens is left alone.
 */
export function cardJson(card: string): string {
  return card.replace(/"(?:[^"\\]|\\.)*"/g, (literal) =>
    literal.replace(/\r/g, '\\r').replace(/\n/g, '\\n'),
  );
}

export function planPostMessage(req: ActuationRequest): PostMessagePlan {
  const text = req.params.text ?? '';
  const card = req.params.html !== undefined ? cardJson(req.params.html) : undefined;
  return {
    text,
    ...(card !== undefined ? { card } : {}),
  };
}
