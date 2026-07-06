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
  /** Gemini Enterprise web/widget config used for user-visible catalogs and widget-style calls. */
  widget?: GeminiWidgetConfig;
  /** Optional model override; otherwise the engine's configured default is used. */
  modelId?: string;
  /**
   * Gemini Enterprise agent/skill resource names to mount on ordinary chat turns.
   * Most deployments should leave this empty; planner/executor skills belong on
   * the explicit route-specific fields below.
   */
  skills?: string[];
  /**
   * Gemini Enterprise widget-style mention markers for ordinary chat skills.
   */
  skillMentions?: GeminiSkillMention[];
  /** Planner skill(s): used only for the pre-execution `plan()` route. */
  plannerSkills?: string[];
  plannerSkillMentions?: GeminiSkillMention[];
  /** Executor skill(s): used only for the constrained command-loop `runCommands()` route. */
  commandSkills?: string[];
  commandSkillMentions?: GeminiSkillMention[];
  /** Data stores selected as default grounding connectors for turns from this pane. */
  dataStores?: string[];
  /**
   * If set, requests are POSTed here instead of directly to discoveryengine.googleapis.com.
   * The only reason to run server code: a tenant that blocks browser CORS or wants a single
   * audited egress point. The proxy is a transparent pass-through (see ADR-0001).
   */
  proxyUrl?: string;
  /** Signed-in user identity (e.g. "v.k@acme") for provenance stamping. */
  identity?: string;
}

export interface GeminiSkillMention {
  label: string;
  uri: string;
}

export type GeminiSkillRoute = 'default' | 'planner' | 'command';

export interface GeminiWidgetConfig {
  configId: string;
  /** Optional widget server token/header if the tenant requires one. Not a Google credential. */
  serverToken?: string;
}

const GLOBAL_HOST = 'https://discoveryengine.googleapis.com';
const CONTENT_GLOBAL_HOST = 'https://content-discoveryengine.googleapis.com';

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

/** Collection resource name (parent of data stores and engines). */
export function collectionResourceName(p: AssistantPath): string {
  const collection = p.collection ?? 'default_collection';
  return `projects/${p.project}/locations/${p.location}/collections/${collection}`;
}

/** Absolute URL for the `:streamAssist` call (or the proxy, if configured). */
export function streamAssistUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/streamAssist`;
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${assistantResourceName(cfg.assistant)}:streamAssist`;
}

/** Absolute URL for the GE widget StreamAssist call used by the hosted GE UI. */
export function widgetStreamAssistUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/widgetStreamAssist`;
  return `${CONTENT_GLOBAL_HOST}/v1alpha/locations/${cfg.assistant.location}/widgetStreamAssist`;
}

/** Absolute URL for GE widget user-visible skill catalog discovery. */
export function widgetListAvailableAgentViewsUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/widgetListAvailableAgentViews`;
  return `${CONTENT_GLOBAL_HOST}/v1alpha/locations/${cfg.assistant.location}/widgetListAvailableAgentViews`;
}

/** Absolute URL for GE widget metadata lookup. */
export function lookupWidgetConfigUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/lookupWidgetConfig`;
  return `${CONTENT_GLOBAL_HOST}/v1alpha/locations/${cfg.assistant.location}/lookupWidgetConfig`;
}

/**
 * Query string for paged catalog list calls: a bounded page size plus the follow-up
 * token from the previous page's `nextPageToken`.
 */
function listPageQuery(pageToken?: string): string {
  const params = new URLSearchParams({ pageSize: '100' });
  if (pageToken) params.set('pageToken', pageToken);
  return `?${params.toString()}`;
}

/** Absolute URL for listing assistant agents/skills (one page; pass `nextPageToken` to follow). */
export function assistantAgentsUrl(cfg: GeminiClientConfig, pageToken?: string): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/catalog/agents${listPageQuery(pageToken)}`;
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${assistantResourceName(cfg.assistant)}/agents${listPageQuery(pageToken)}`;
}

/** Absolute URL for listing collection data stores (one page; pass `nextPageToken` to follow). */
export function dataStoresUrl(cfg: GeminiClientConfig, pageToken?: string): string {
  if (cfg.proxyUrl) {
    return `${proxyBase(cfg.proxyUrl)}/catalog/dataStores${listPageQuery(pageToken)}`;
  }
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${collectionResourceName(cfg.assistant)}/dataStores${listPageQuery(pageToken)}`;
}

/**
 * Absolute URL for listing project+location collections (one page; pass `nextPageToken` to
 * follow). Each collection embeds its output-only `dataConnector`, so this single list is the
 * whole connector catalog — no per-connector GET is needed.
 */
export function collectionsUrl(cfg: GeminiClientConfig, pageToken?: string): string {
  if (cfg.proxyUrl) {
    return `${proxyBase(cfg.proxyUrl)}/catalog/collections${listPageQuery(pageToken)}`;
  }
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${projectLocationResourceName(cfg.assistant)}/collections${listPageQuery(pageToken)}`;
}

/** Absolute URL for listing engine sessions. */
export function sessionsUrl(cfg: GeminiClientConfig): string {
  if (cfg.proxyUrl) return `${proxyBase(cfg.proxyUrl)}/sessions`;
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${engineResourceName(cfg.assistant)}/sessions`;
}

/** Absolute URL for getting one engine session. */
export function sessionUrl(cfg: GeminiClientConfig, sessionIdOrName: string): string {
  if (sessionIdOrName.startsWith('projects/')) {
    const host = cfg.proxyUrl
      ? proxyBase(cfg.proxyUrl)
      : discoveryEngineHost(cfg.assistant.location);
    return `${host}/v1alpha/${sessionIdOrName}`;
  }
  return `${sessionsUrl(cfg)}/${encodeURIComponent(sessionIdOrName)}`;
}

/** Absolute URL for adding a session context file (`v1` addContextFile). */
export function addContextFileUrl(cfg: GeminiClientConfig, sessionIdOrName: string): string {
  if (cfg.proxyUrl) {
    return `${proxyBase(cfg.proxyUrl)}/sessions/${encodeURIComponent(sessionIdOrName)}:addContextFile`;
  }
  const host = discoveryEngineHost(cfg.assistant.location);
  if (sessionIdOrName.startsWith('projects/'))
    return `${host}/v1/${sessionIdOrName}:addContextFile`;
  return `${host}/v1/${engineResourceName(cfg.assistant)}/sessions/${encodeURIComponent(
    sessionIdOrName,
  )}:addContextFile`;
}

/** Absolute URL for listing metadata for context files in one session. */
export function sessionFilesUrl(cfg: GeminiClientConfig, sessionIdOrName: string): string {
  if (cfg.proxyUrl) {
    return `${proxyBase(cfg.proxyUrl)}/sessions/${encodeURIComponent(sessionIdOrName)}/files`;
  }
  const host = discoveryEngineHost(cfg.assistant.location);
  if (sessionIdOrName.startsWith('projects/')) return `${host}/v1/${sessionIdOrName}/files`;
  return `${host}/v1/${engineResourceName(cfg.assistant)}/sessions/${encodeURIComponent(
    sessionIdOrName,
  )}/files`;
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
