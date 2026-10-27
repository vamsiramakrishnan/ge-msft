/**
 * Microsoft Graph endpoint + the **delegated** scopes the add-in requests. Everything here
 * acts as the signed-in user (NAA delegated token), never an app principal — the add-in's
 * blast radius equals the user's own access. Prefer the narrowest scope (e.g. Sites.Selected
 * over Sites.Read.All) in the manifest; this is the read-side set.
 */
export interface GraphConfig {
  /** Override for sovereign/regional clouds; default global v1.0. */
  baseUrl?: string;
  /**
   * Hosts a chat attachment's `contentUrl` may point at before it is resolved as the user. A chat
   * member can post a reference to any URL; only the tenant's SharePoint/OneDrive hosts are
   * resolved, so a foreign link never becomes a listed "shared file". Suffix match on the host
   * (`sharepoint.com` covers `contoso.sharepoint.com` and `contoso-my.sharepoint.com`). Override
   * for sovereign clouds (`sharepoint.us`, `sharepoint.de`, `sharepoint.cn`).
   */
  sharingHosts?: readonly string[];
}

export const DEFAULT_SHARING_HOSTS: readonly string[] = ['sharepoint.com'];

/** True when `url` is an https link whose host is (a subdomain of) an allowed sharing host. */
export function isAllowedSharingUrl(
  url: string,
  hosts: readonly string[] = DEFAULT_SHARING_HOSTS,
): boolean {
  let parsed: URL;
  try {
    parsed = new URL(url);
  } catch {
    return false;
  }
  if (parsed.protocol !== 'https:') return false;
  const host = parsed.hostname.toLowerCase();
  return hosts.some((h) => {
    const suffix = h.toLowerCase();
    return host === suffix || host.endsWith(`.${suffix}`);
  });
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
  /**
   * The cross-surface handoff store (`/shared` DocFs mount). `Files.ReadWrite.AppFolder` is
   * deliberately the NARROWEST write scope Graph offers — it grants read/write ONLY to a special
   * per-app folder (`/me/drive/special/approot`) that other apps (and the user's own OneDrive
   * browsing UI, by default) cannot see, never the broader `files`/`search` read scopes above. This
   * is a NEW consent prompt beyond what this add-in previously requested — flag it as such
   * wherever scopes are surfaced to the user/admin, and route any use of it through a security
   * review (identity/credentials territory) before shipping.
   */
  shared: ['Files.ReadWrite.AppFolder'],
  /**
   * Files shared in the current Teams chat or meeting: the chat's messages name the files
   * (`attachments[].contentUrl`), and the file itself is read with the `files` scopes as the user.
   * `Chat.Read` is delegated and user-scoped: only chats the signed-in user is a member of.
   */
  chat: ['Chat.Read'],
} as const;

export function graphUrl(cfg: GraphConfig, path: string): string {
  const base = (cfg.baseUrl ?? GRAPH_BASE_URL).replace(/\/$/, '');
  return path.startsWith('http') ? path : `${base}${path.startsWith('/') ? '' : '/'}${path}`;
}
