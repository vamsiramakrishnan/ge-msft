/**
 * @ge/graph-client — the add-in's own Microsoft Graph reader (Plane B). Reads mail,
 * calendar, files, people, and runs Microsoft Search **as the signed-in user** (delegated
 * NAA token, no secrets) and resolves each into the shared `ResolvedContext`, so estate
 * context attaches to a session exactly like in-document context. The add-in does this
 * itself — not Gemini Enterprise. See docs/ACCESS-MODEL.md (Plane B).
 */
export * from './config.js';
export {
  GraphClient,
  messageToContext,
  eventToContext,
  userToContext,
  driveItemToContext,
  type GraphTokenSource,
} from './graph-client.js';
export type { GraphMessage, GraphEvent, GraphDriveItem, GraphUser } from './graph-types.js';
