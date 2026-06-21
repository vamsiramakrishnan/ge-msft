/**
 * @ge/bridge-outlook — the Outlook DocBridge. String-path context capture (the mail item's
 * subject / from / body, labelled and normalized via @ge/content) + grounded reply/compose
 * writes through `Office.context.mailbox.item`. Implements @ge/runtime's DocBridge, so it
 * plugs into the shared AssistSession loop unchanged.
 */
export { OutlookBridge } from './outlook-bridge.js';
export { OUTLOOK_CAPABILITIES } from './capabilities.js';
export { mailItemToContext, type MailItem } from './capture.js';
export { planReply, type ReplyPlan } from './actuate-plan.js';
export { composeEvent, receivedEvent, sendEvent, decideSend, type SendDecision } from './events.js';
export {
  createMessageSendHandler,
  activeItemIdResolver,
  type MessageSendHandler,
  type MessageSendHandlerOptions,
  type OnSendEvent,
  type OnSendCompletedOptions,
} from './on-send.js';
