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

/**
 * Regional host for residency (e.g. 'eu', 'us', 'asia-northeast1'). Only the explicit value
 * `'global'` selects the global host — a missing/empty location is a configuration error, never
 * a silent global fallback, so a residency-pinned tenant cannot be downgraded by misconfig.
 */
export function discoveryEngineHost(location: string): string {
  if (location === 'global') return GLOBAL_HOST;
  if (!location) {
    throw new Error(
      'Discovery Engine location is required (residency pin): set assistant.location to a ' +
        'region (e.g. "eu", "us") or the explicit value "global".',
    );
  }
  return `https://discoveryengine.${location}.rep.googleapis.com`;
}

/**
 * Validate + normalize the optional egress proxy. The federated bearer token is attached to this
 * origin (see de-fetch.ts), so it MUST be HTTPS — otherwise the token could leak in cleartext or
 * to an unintended origin. `localhost`/`127.0.0.1` may use HTTP for local dev only.
 */
export function proxyBase(proxyUrl: string): string {
  let parsed: URL;
  try {
    parsed = new URL(proxyUrl);
  } catch {
    throw new Error(`Invalid proxyUrl: ${proxyUrl}`);
  }
  const isLocalhost = parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
  if (parsed.protocol !== 'https:' && !(parsed.protocol === 'http:' && isLocalhost)) {
    throw new Error(
      `proxyUrl must be https (the federated token is attached to it): got ${parsed.protocol}`,
    );
  }
  return proxyUrl.replace(/\/$/, '');
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
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/streamAssist`;
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${assistantResourceName(cfg.assistant)}:streamAssist`;
}

/** Engine resource name (parent of sessions/servingConfigs), for session creation etc. */
export function engineResourceName(p: AssistantPath): string {
  const collection = p.collection ?? 'default_collection';
  return `projects/${p.project}/locations/${p.location}/collections/${collection}/engines/${p.engine}`;
}

/** Project+location resource name (parent of grounding/ranking configs). */
export function projectLocationResourceName(p: AssistantPath): string {
  return `projects/${p.project}/locations/${p.location}`;
}

/**
 * Absolute URL for the engine-scoped `:search` serving config call (or the proxy).
 * Default serving config id is `default_search`.
 */
export function searchUrl(cfg: GeminiClientConfig, servingConfig = 'default_search'): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/search`;
  const host = discoveryEngineHost(cfg.assistant.location);
  const engine = engineResourceName(cfg.assistant);
  return `${host}/v1alpha/${engine}/servingConfigs/${servingConfig}:search`;
}

/** Absolute URL for the engine-scoped `completionConfig:completeQuery` call (or the proxy). */
export function completeQueryUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/completeQuery`;
  const host = discoveryEngineHost(cfg.assistant.location);
  const engine = engineResourceName(cfg.assistant);
  return `${host}/v1alpha/${engine}/completionConfig:completeQuery`;
}

/**
 * Absolute URL for the project+location-scoped `groundingConfigs:check` call (or the proxy).
 * Default grounding config id is `default_grounding_config`. NOT under the engine.
 */
export function checkGroundingUrl(
  cfg: GeminiClientConfig,
  groundingConfig = 'default_grounding_config',
): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/checkGrounding`;
  const host = discoveryEngineHost(cfg.assistant.location);
  const base = projectLocationResourceName(cfg.assistant);
  return `${host}/v1alpha/${base}/groundingConfigs/${groundingConfig}:check`;
}

/**
 * Absolute URL for the project+location-scoped `rankingConfigs:rank` call (or the proxy).
 * Default ranking config id is `default_ranking_config`. NOT under the engine.
 */
export function rankUrl(cfg: GeminiClientConfig, rankingConfig = 'default_ranking_config'): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/rank`;
  const host = discoveryEngineHost(cfg.assistant.location);
  const base = projectLocationResourceName(cfg.assistant);
  return `${host}/v1alpha/${base}/rankingConfigs/${rankingConfig}:rank`;
}
