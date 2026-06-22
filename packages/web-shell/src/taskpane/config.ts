import type { AssistantPath, WifConfig } from '@ge/gemini-client';
import type { ShellConfig } from '../compose.js';
import type { NaaAuthOptions } from '../auth-client.js';

/**
 * Read-only, typed view over the Vite-injected environment (`import.meta.env`). Every value the
 * shell needs to reach Gemini Enterprise as the signed-in user is sourced here — never hardcoded.
 * No Google secret is ever read or held: only the GCP project/engine coordinates and the WIF
 * pool/provider the user's Entra token federates to. See `.env.example`.
 */
export interface RawEnv {
  readonly VITE_GCP_PROJECT?: string;
  readonly VITE_GCP_LOCATION?: string;
  readonly VITE_GE_ENGINE?: string;
  readonly VITE_GE_COLLECTION?: string;
  readonly VITE_GE_ASSISTANT?: string;
  readonly VITE_GE_MODEL_ID?: string;
  readonly VITE_WIF_POOL_ID?: string;
  readonly VITE_WIF_PROVIDER_ID?: string;
  readonly VITE_WIF_SCOPE?: string;
  readonly VITE_WIF_USER_PROJECT?: string;
  readonly VITE_PROXY_URL?: string;
  readonly VITE_ENTRA_CLIENT_ID?: string;
  readonly VITE_ENTRA_AUTHORITY?: string;
  readonly VITE_WIF_ID_TOKEN_SCOPES?: string;
  readonly VITE_GRAPH_SCOPES?: string;
  readonly VITE_NOTEBOOK_ID?: string;
}

function required(env: RawEnv, key: keyof RawEnv): string {
  const value = env[key];
  if (value === undefined || value === '') {
    throw new Error(
      `Missing required environment variable ${key}. Copy .env.example to .env and fill it in.`,
    );
  }
  return value;
}

function splitScopes(value: string | undefined, fallback: string[]): string[] {
  if (!value) return fallback;
  const scopes = value
    .split(/[, ]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return scopes.length ? scopes : fallback;
}

/** The Discovery Engine assistant coordinates (residency pinned by `location`). */
export function assistantFromEnv(env: RawEnv): AssistantPath {
  return {
    project: required(env, 'VITE_GCP_PROJECT'),
    location: required(env, 'VITE_GCP_LOCATION'),
    engine: required(env, 'VITE_GE_ENGINE'),
    ...(env.VITE_GE_COLLECTION ? { collection: env.VITE_GE_COLLECTION } : {}),
    ...(env.VITE_GE_ASSISTANT ? { assistant: env.VITE_GE_ASSISTANT } : {}),
  };
}

/** The Workforce Identity Federation pool/provider the Entra id token exchanges against. */
export function wifFromEnv(env: RawEnv): WifConfig {
  return {
    poolId: required(env, 'VITE_WIF_POOL_ID'),
    providerId: required(env, 'VITE_WIF_PROVIDER_ID'),
    ...(env.VITE_WIF_SCOPE ? { scope: env.VITE_WIF_SCOPE } : {}),
    ...(env.VITE_WIF_USER_PROJECT ? { userProject: env.VITE_WIF_USER_PROJECT } : {}),
  };
}

/** The full shell config: assistant + WIF + optional model/proxy. No credentials. */
export function shellConfigFromEnv(env: RawEnv): ShellConfig {
  return {
    assistant: assistantFromEnv(env),
    wif: wifFromEnv(env),
    ...(env.VITE_GE_MODEL_ID ? { modelId: env.VITE_GE_MODEL_ID } : {}),
    ...(env.VITE_PROXY_URL ? { proxyUrl: env.VITE_PROXY_URL } : {}),
  };
}

/** MSAL/NAA scopes: the id-token (WIF subject) scopes and default Graph scopes. */
export function authOptionsFromEnv(env: RawEnv): NaaAuthOptions {
  const idTokenScopes = splitScopes(env.VITE_WIF_ID_TOKEN_SCOPES, [
    `${env.VITE_ENTRA_CLIENT_ID ?? ''}/.default`.replace(/^\/\.default$/, 'openid'),
  ]);
  // Least-privilege default: only the signed-in user's own profile. Plane-B estate reads add
  // narrower, source-scoped delegated permissions (prefer `Sites.Selected` over `*.All`) via
  // VITE_GRAPH_SCOPES per deployment — never default to a tenant-wide `*.All` read here.
  const graphScopes = splitScopes(env.VITE_GRAPH_SCOPES, ['User.Read']);
  return { idTokenScopes, graphScopes };
}

export interface MsalConfigLike {
  clientId: string;
  authority?: string;
}

/** The Entra app coordinates for MSAL NAA bootstrap (no secret — public client / NAA). */
export function msalConfigFromEnv(env: RawEnv): MsalConfigLike {
  return {
    clientId: required(env, 'VITE_ENTRA_CLIENT_ID'),
    ...(env.VITE_ENTRA_AUTHORITY ? { authority: env.VITE_ENTRA_AUTHORITY } : {}),
  };
}

/** Optional notebook id that seeds the research unit's curated core. */
export function notebookIdFromEnv(env: RawEnv): string | undefined {
  return env.VITE_NOTEBOOK_ID || undefined;
}
