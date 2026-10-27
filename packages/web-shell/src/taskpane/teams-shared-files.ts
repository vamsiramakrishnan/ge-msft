import type { SharedFile, SharedFilesSource } from '@ge/teams';
import { DEFAULT_CONTEXT_FILE_MAX_BYTES } from '@ge/gemini-client';
import type { GraphClient } from '@ge/graph-client';

/**
 * Files shared in the current Teams chat or meeting, read as the signed-in user through Microsoft
 * Graph (Plane B). This is the shell-side adapter the Teams bridge is handed via `setSharedFiles`:
 * the bridge never calls Graph itself. Listing uses the chat's messages (`Chat.Read`, delegated,
 * the user's own chats only); reading uses the user's file permissions, so a file the user cannot
 * open is simply absent or fails to read. Nothing is cached across sessions.
 */

/** The TeamsJS context slice that names the current chat (a meeting side panel carries its chat). */
export interface TeamsContextLike {
  chat?: { id?: string };
  meeting?: { id?: string };
}

/** Resolve the chat id for the current Teams frame via `app.getContext()`, or none. */
export async function currentTeamsChatId(
  teams: { app?: { getContext?: () => Promise<TeamsContextLike> } } | undefined,
): Promise<string | undefined> {
  const getContext = teams?.app?.getContext;
  if (typeof getContext !== 'function') return undefined;
  try {
    const ctx = await getContext();
    return ctx?.chat?.id || undefined;
  } catch {
    return undefined;
  }
}

export function graphSharedFiles(
  graph: Pick<GraphClient, 'listChatFiles' | 'getDriveItemContent'>,
  chatId: string,
  opts: { maxBytes?: number } = {},
): SharedFilesSource {
  const maxBytes = opts.maxBytes ?? DEFAULT_CONTEXT_FILE_MAX_BYTES;
  let known = new Map<string, { driveId?: string }>();
  return {
    async list(): Promise<SharedFile[]> {
      const files = await graph.listChatFiles(chatId);
      known = new Map(files.map((f) => [f.id, { ...(f.driveId ? { driveId: f.driveId } : {}) }]));
      return files.map((f) => ({
        id: f.id,
        name: f.name,
        ...(f.mimeType ? { mimeType: f.mimeType } : {}),
        ...(f.sizeBytes !== undefined ? { sizeBytes: f.sizeBytes } : {}),
        ...(f.webUrl ? { webUrl: f.webUrl } : {}),
        ...(f.sharedBy ? { sharedBy: f.sharedBy } : {}),
      }));
    },
    async read(id: string) {
      const entry = known.get(id);
      if (!entry) throw new Error('This file was not listed in the current chat.');
      return graph.getDriveItemContent(id, entry.driveId, { maxBytes });
    },
  };
}
