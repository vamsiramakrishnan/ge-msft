/**
 * Endpoint + resource-path construction for Discovery Engine (Gemini Enterprise).
 * The add-in calls these endpoints directly as the signed-in user; residency is
 * pinned by choosing the regional endpoint.
 */

/** Fully-qualifies an assistant resource: the engine + assistant the add-in talks to. */
export interface AssistantPath {
  project: string; // GCP project id or number
  location: string; // 'global' | 'us' | 'eu' | regional id — match residency
  collection?: string; // default: 'default_collection'
  engine: string; // the Gemini Enterprise app/engine id
  assistant?: string; // default: 'default_assistant'
}

export interface GeminiClientConfig {
  assistant: AssistantPath;
  /** Optional model override; otherwise the engine's configured default is used. */
  modelId?: string;
  /**
   * If set, requests are POSTed here instead of directly to discoveryengine.googleapis.com.
   * The only reason to run server code: a tenant that blocks browser CORS or wants a single
   * audited egress point. The proxy is a transparent pass-through (see ADR-0001).
   */
  proxyUrl?: string;
  /** Signed-in user identity (e.g. "v.k@acme") for provenance stamping. */
  identity?: string;
}

const GLOBAL_HOST = 'https://discoveryengine.googleapis.com';

/** Regional host for residency (e.g. 'eu', 'us', 'asia-northeast1'); 'global' → global host. */
export function discoveryEngineHost(location: string): string {
  if (!location || location === 'global') return GLOBAL_HOST;
  return `https://discoveryengine.${location}.rep.googleapis.com`;
}

export function assistantResourceName(p: AssistantPath): string {
  const collection = p.collection ?? 'default_collection';
  const assistant = p.assistant ?? 'default_assistant';
  return (
    `projects/${p.project}/locations/${p.location}/collections/${collection}` +
    `/engines/${p.engine}/assistants/${assistant}`
  );
}

/** Absolute URL for the `:streamAssist` call (or the proxy, if configured). */
export function streamAssistUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${cfg.proxyUrl.replace(/\/$/, '')}/streamAssist`;
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${assistantResourceName(cfg.assistant)}:streamAssist`;
}

/** Engine resource name (parent of sessions/servingConfigs), for session creation etc. */
export function engineResourceName(p: AssistantPath): string {
  const collection = p.collection ?? 'default_collection';
  return `projects/${p.project}/locations/${p.location}/collections/${collection}/engines/${p.engine}`;
}
