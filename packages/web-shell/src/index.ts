/**
 * @ge/web-shell — the surface-agnostic web app core. It turns the signed-in user's identity +
 * a surface bridge into a running, grounded assist session with a context tray, streaming chat,
 * and reversible actuation review. Host pixels (React) and Office bootstrapping sit one layer up;
 * everything here is framework-free and unit-tested.
 */
export { surfaceFromHost, detectSurface, type OfficeContextLike } from './host.js';
export {
  NaaAuthClient,
  type MsalLike,
  type MsalAccount,
  type MsalAuthResult,
  type MsalTokenRequest,
  type NaaAuthOptions,
} from './auth-client.js';
export { ProvenanceStore, type ChangeRecord } from './provenance-store.js';
export {
  composeSession,
  type ShellConfig,
  type ComposeOptions,
  type ComposedSession,
} from './compose.js';
export {
  PanelController,
  type AssistLike,
  type ContextLister,
  type PanelState,
  type ChatMessage,
  type ContextChip,
  type Suggestion,
  type Proposal,
} from './controller.js';
