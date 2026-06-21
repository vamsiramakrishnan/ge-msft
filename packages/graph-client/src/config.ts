/**
 * Microsoft Graph endpoint + the **delegated** scopes the add-in requests. Everything here
 * acts as the signed-in user (NAA delegated token), never an app principal — the add-in's
 * blast radius equals the user's own access. Prefer the narrowest scope (e.g. Sites.Selected
 * over Sites.Read.All) in the manifest; this is the read-side set.
 */
export interface GraphConfig {
  /** Override for sovereign/regional clouds; default global v1.0. */
  baseUrl?: string;
}

export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/** Delegated read scopes by capability (declared in the unified manifest, user/admin-consented). */
export const GRAPH_SCOPES = {
  basic: ['User.Read', 'openid', 'profile', 'offline_access'],
  mail: ['Mail.Read'],
  calendar: ['Calendars.Read'],
  files: ['Files.Read.All', 'Sites.Read.All'],
  people: ['User.ReadBasic.All'],
  search: ['Files.Read.All', 'Sites.Read.All', 'Mail.Read'],
} as const;

export function graphUrl(cfg: GraphConfig, path: string): string {
  const base = (cfg.baseUrl ?? GRAPH_BASE_URL).replace(/\/$/, '');
  return path.startsWith('http') ? path : `${base}${path.startsWith('/') ? '' : '/'}${path}`;
}
