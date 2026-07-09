import type { EstateRef, ResolvedContext } from '@ge/contracts';
import { toContext } from '@ge/content';
import { GraphConfig, GRAPH_SCOPES, graphUrl } from './config.js';
import {
  GraphChildrenSchema,
  GraphDriveItemSchema,
  GraphEventSchema,
  GraphMessageSchema,
  GraphSearchResponseSchema,
  GraphUserSchema,
  type GraphDriveItem,
  type GraphEvent,
  type GraphMessage,
  type GraphUser,
} from './graph-types.js';
import { z } from 'zod';

/** Supplies a delegated Graph access token for the given scopes (e.g. the AuthClient via NAA). */
export interface GraphTokenSource {
  getGraphToken(scopes: string[]): Promise<string>;
  invalidate?(): void;
}

type FetchLike = typeof fetch;

const defaultFetch: FetchLike = (input, init) => globalThis.fetch(input, init);

/**
 * Reads the user's Microsoft 365 estate (Plane B) **directly from the add-in**, as the
 * signed-in user, and turns each item into the same `ResolvedContext` the rest of the
 * pipeline consumes. No service-account keys; only the user's delegated Graph token.
 */
export class GraphClient {
  constructor(
    private readonly tokens: GraphTokenSource,
    private readonly config: GraphConfig = {},
    private readonly fetchImpl: FetchLike = defaultFetch,
  ) {}

  // ---- typed reads ----

  async getMessage(id: string): Promise<GraphMessage> {
    return this.getJson(`/me/messages/${enc(id)}`, [...GRAPH_SCOPES.mail], GraphMessageSchema);
  }

  async getEvent(id: string): Promise<GraphEvent> {
    return this.getJson(`/me/events/${enc(id)}`, [...GRAPH_SCOPES.calendar], GraphEventSchema);
  }

  async getDriveItem(id: string, driveId?: string): Promise<GraphDriveItem> {
    const path = driveId
      ? `/drives/${enc(driveId)}/items/${enc(id)}`
      : `/me/drive/items/${enc(id)}`;
    return this.getJson(path, [...GRAPH_SCOPES.files], GraphDriveItemSchema);
  }

  async getUser(id: string): Promise<GraphUser> {
    return this.getJson(`/users/${enc(id)}`, [...GRAPH_SCOPES.people], GraphUserSchema);
  }

  /**
   * Microsoft Search across the estate as the user → lightweight `EstateRef`s the UI can
   * offer to attach (the brilliant bit: search SharePoint/OneDrive/mail without leaving the pane).
   */
  async search(
    queryString: string,
    entityTypes: Array<'driveItem' | 'listItem' | 'message' | 'event'> = ['driveItem', 'listItem'],
    size = 10,
  ): Promise<EstateRef[]> {
    const body = {
      requests: [{ entityTypes, query: { queryString }, from: 0, size }],
    };
    const res = await this.post(
      '/search/query',
      body,
      [...GRAPH_SCOPES.search],
      GraphSearchResponseSchema,
    );
    const refs: EstateRef[] = [];
    for (const r of res.value ?? []) {
      for (const c of r.hitsContainers ?? []) {
        for (const hit of c.hits ?? []) {
          const ref = hitToEstateRef(hit.hitId, hit.summary, hit.resource);
          if (ref) refs.push(ref);
        }
      }
    }
    return refs;
  }

  // ---- resolve any EstateRef into attach-ready context ----

  async resolveEstateRef(ref: EstateRef): Promise<ResolvedContext[]> {
    switch (ref.source) {
      case 'mail':
      case 'mail-thread':
        return messageToContext(await this.getMessage(ref.id));
      case 'calendar':
        return eventToContext(await this.getEvent(ref.id));
      case 'drive-item':
        return driveItemToContext(await this.getDriveItem(ref.id, ref.driveId), ref);
      case 'site':
        // List/site items resolve as a titled reference back to the source.
        return referenceContext(ref);
      case 'person':
        return userToContext(await this.getUser(ref.id));
    }
  }

  // ---- /shared: the cross-surface handoff store (Files.ReadWrite.AppFolder, see config.ts) ----
  // A per-app OneDrive folder (`/me/drive/special/approot`) invisible to other apps and not shown
  // in the user's normal OneDrive browsing UI by default — the narrowest available Graph write
  // surface. This is what the `/shared` DocFs mount is backed by (see docfs/shared-mount.ts).

  /** List files at the top level of the app folder — `{name, size}` per file, newest-Graph-order. */
  async listSharedFiles(): Promise<{ name: string; size: number }[]> {
    const res = await this.getJson(
      '/me/drive/special/approot/children',
      [...GRAPH_SCOPES.shared],
      GraphChildrenSchema,
    );
    return (res.value ?? [])
      .filter((item) => item.file !== undefined)
      .map((item) => ({ name: item.name, size: item.size ?? 0 }));
  }

  /** Read one app-folder file's text content, or `undefined` if it doesn't exist (404 → undefined,
   *  not a throw — a missing shared file is an expected, everyday case, not an error). */
  async getSharedFile(path: string): Promise<string | undefined> {
    const res = await this.send('GET', `/me/drive/special/approot:/${encPath(path)}:/content`, [
      ...GRAPH_SCOPES.shared,
    ]).catch((err: unknown) => {
      if (err instanceof GraphNotFoundError) return undefined;
      throw err;
    });
    return res ? res.text() : undefined;
  }

  /** Write (create or overwrite) one app-folder file's text content. */
  async putSharedFile(path: string, content: string): Promise<void> {
    await this.send(
      'PUT',
      `/me/drive/special/approot:/${encPath(path)}:/content`,
      [...GRAPH_SCOPES.shared],
      content,
      'text/plain',
    );
  }

  /** Delete one app-folder file. A missing file is treated as already-deleted (404 is not an error). */
  async deleteSharedFile(path: string): Promise<void> {
    await this.send('DELETE', `/me/drive/special/approot:/${encPath(path)}`, [
      ...GRAPH_SCOPES.shared,
    ]).catch((err: unknown) => {
      if (err instanceof GraphNotFoundError) return undefined;
      throw err;
    });
  }

  // ---- transport ----

  private async getJson<T>(path: string, scopes: string[], schema: z.ZodType<T>): Promise<T> {
    const res = await this.send('GET', path, scopes);
    return schema.parse(await res.json());
  }

  private async post<T>(
    path: string,
    body: unknown,
    scopes: string[],
    schema: z.ZodType<T>,
  ): Promise<T> {
    const res = await this.send('POST', path, scopes, JSON.stringify(body));
    return schema.parse(await res.json());
  }

  private async send(
    method: 'GET' | 'POST' | 'PUT' | 'DELETE',
    path: string,
    scopes: string[],
    body?: string,
    contentType: string = 'application/json',
  ): Promise<Response> {
    const url = graphUrl(this.config, path);
    const call = async (): Promise<Response> => {
      const token = await this.tokens.getGraphToken(scopes);
      const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
      if (body) headers['Content-Type'] = contentType;
      return this.fetchImpl(url, { method, headers, ...(body ? { body } : {}) });
    };
    let res = await call();
    if (res.status === 401 && this.tokens.invalidate) {
      this.tokens.invalidate();
      res = await call();
    }
    if (res.status === 404) {
      throw new GraphNotFoundError(`Graph ${method} ${path} — not found`);
    }
    if (!res.ok) {
      throw new Error(`Graph ${method} ${path} failed (${res.status}): ${await safeText(res)}`);
    }
    return res;
  }
}

/** Thrown by `send()` on a Graph 404 — callers that treat "missing" as expected (not exceptional)
 *  catch this specifically rather than string-matching an error message. */
export class GraphNotFoundError extends Error {}

// ---- mappers (pure-ish; only depend on the parsed Graph shapes) ----

export function messageToContext(m: GraphMessage): ResolvedContext[] {
  const from = m.from?.emailAddress;
  const header = [
    m.subject ? `Subject: ${m.subject}` : '',
    from ? `From: ${from.name ?? ''} <${from.address ?? ''}>` : '',
    m.receivedDateTime ? `Date: ${m.receivedDateTime}` : '',
  ]
    .filter(Boolean)
    .join('\n');
  const isHtml = (m.body?.contentType ?? '').toLowerCase() === 'html';
  const text = `${header}\n\n${m.body?.content ?? m.bodyPreview ?? ''}`;
  return toContext({
    sourceId: `graph:mail:${m.id ?? 'item'}`,
    text,
    format: isHtml ? 'html' : 'plain',
    title: m.subject ?? 'Email',
    surface: 'outlook',
  });
}

export function eventToContext(e: GraphEvent): ResolvedContext[] {
  const header = [
    e.subject ? `Event: ${e.subject}` : '',
    e.start?.dateTime ? `Start: ${e.start.dateTime}` : '',
    e.end?.dateTime ? `End: ${e.end.dateTime}` : '',
    e.location?.displayName ? `Location: ${e.location.displayName}` : '',
  ]
    .filter(Boolean)
    .join('\n');
  const text = `${header}\n\n${e.body?.content ?? e.bodyPreview ?? ''}`;
  return toContext({
    sourceId: `graph:event:${e.id ?? 'item'}`,
    text,
    format: (e.body?.contentType ?? '').toLowerCase() === 'html' ? 'html' : 'plain',
    title: e.subject ?? 'Calendar event',
    surface: 'outlook',
  });
}

export function userToContext(u: GraphUser): ResolvedContext[] {
  const text = [
    u.displayName ? `Name: ${u.displayName}` : '',
    u.jobTitle ? `Title: ${u.jobTitle}` : '',
    (u.mail ?? u.userPrincipalName) ? `Email: ${u.mail ?? u.userPrincipalName}` : '',
  ]
    .filter(Boolean)
    .join('\n');
  return toContext({
    sourceId: `graph:person:${u.id ?? 'user'}`,
    text,
    format: 'plain',
    title: u.displayName ?? 'Person',
  });
}

/**
 * A file resolves to a titled, deep-linked reference (we don't inline binary content) — the
 * reference-over-inline policy. The pane can offer "open" via webUrl; the engine can ground
 * on the indexed copy if the file is in a connected data store.
 */
export function driveItemToContext(item: GraphDriveItem, ref: EstateRef): ResolvedContext[] {
  const title = item.name ?? ref.title ?? 'File';
  return [
    {
      ref: {
        id: `graph:drive:${item.id ?? ref.id}`,
        kind: 'drive-document',
        surface: 'word',
        title,
        ...(item.webUrl ? { preview: item.webUrl } : {}),
        ...(item.size ? { sizeBytes: item.size } : {}),
      },
      value: {
        as: 'drive-document',
        driveId: item.parentReference?.driveId ?? ref.driveId ?? '',
        documentName: item.id ?? ref.id,
        title,
      },
    },
  ];
}

function referenceContext(ref: EstateRef): ResolvedContext[] {
  const title = ref.title ?? 'Item';
  return [
    {
      ref: {
        id: `graph:${ref.source}:${ref.id}`,
        kind: 'indexed-document',
        surface: 'word',
        title,
        ...(ref.webUrl ? { preview: ref.webUrl } : {}),
      },
      value: { as: 'indexed-document', documentName: ref.id, title },
    },
  ];
}

function hitToEstateRef(
  hitId: string | undefined,
  summary: string | undefined,
  resource: Record<string, unknown> | undefined,
): EstateRef | undefined {
  if (!resource) return undefined;
  const odata = String(resource['@odata.type'] ?? '').toLowerCase();
  const id = String(resource['id'] ?? hitId ?? '');
  if (!id) return undefined;
  const name = (resource['name'] ?? resource['subject'] ?? resource['displayName']) as
    | string
    | undefined;
  const webUrl = (resource['webUrl'] ?? resource['webLink']) as string | undefined;
  const base = {
    id,
    ...(name ? { title: name } : {}),
    ...(summary ? { preview: summary } : {}),
    ...(webUrl ? { webUrl } : {}),
  };
  if (odata.includes('driveitem')) {
    const driveId = (resource['parentReference'] as { driveId?: string } | undefined)?.driveId;
    return { source: 'drive-item', ...base, ...(driveId ? { driveId } : {}) };
  }
  if (odata.includes('message')) return { source: 'mail', ...base };
  if (odata.includes('event')) return { source: 'calendar', ...base };
  if (odata.includes('listitem')) return { source: 'site', ...base };
  return { source: 'drive-item', ...base };
}

function enc(s: string): string {
  return encodeURIComponent(s);
}

/**
 * Encode a `/`-separated relative path for Graph's `:/path:/content` addressing — each segment is
 * percent-encoded individually so a literal `/` stays a path separator, not `%2F`.
 *
 * Defense-in-depth: `packages/contracts`'s `normalizeWorkspaceName` already rejects `..`/leading-
 * or-trailing-`/`/empty names before a `share` command ever compiles, but that check lives one
 * layer up (the command grammar), not here. Re-validating at the point of the actual Graph call
 * means every current AND future caller of `getSharedFile`/`putSharedFile`/`deleteSharedFile` is
 * protected from addressing outside the app folder (`/me/drive/special/approot:/../x` etc.),
 * not just the one path the grammar happens to gate today.
 */
function encPath(path: string): string {
  const segments = path.split('/');
  for (const segment of segments) {
    if (segment === '' || segment === '.' || segment === '..') {
      throw new Error(
        `invalid /shared path "${path}" — must not contain empty, ".", or ".." segments`,
      );
    }
  }
  return segments.map((segment) => enc(segment)).join('/');
}

async function safeText(res: Response): Promise<string> {
  try {
    return (await res.text()).slice(0, 300);
  } catch {
    return '';
  }
}
