import { z } from 'zod';
import {
  dataStoreResourceName,
  dataStoreUrl,
  engineUrl,
  type GeminiClientConfig,
} from './config.js';
import { defaultFetch, getJson, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/**
 * Federated data-store discovery + selection, driven by the signed-in user's WIF token.
 *
 * Verified live (saib, 2026-07): `dataStores.list` / `collections.list` are 403 for the WIF principal,
 * but **`engines.get` is permitted and the Engine resource carries `dataStoreIds[]`** — the stores
 * attached to the app (Outlook / OneDrive / SharePoint federated connectors). So we discover the
 * groundable data stores via `engines.get` (not `dataStores.list`), read each with `dataStores.get`,
 * and pass the chosen ids into `streamAssist` `toolsSpec.vertexAiSearchSpec.dataStoreSpecs[]`.
 * See docs/api/discoveryengine/listing-agents-and-skills.md.
 */

export const EngineSchema = z.object({
  name: z.string().optional(),
  displayName: z.string().optional(),
  dataStoreIds: z.array(z.string()).optional(),
});

export const DataStoreSchema = z.object({
  name: z.string().optional(),
  displayName: z.string().optional(),
  industryVertical: z.string().optional(),
});

export interface EngineDataStore {
  /** Bare data store id (the `dataStoreIds[]` entry), e.g. `msft-sharepoint-fed_…_file`. */
  id: string;
  /** Full resource name for `dataStoreSpecs[].dataStore`. */
  resourceName: string;
  /** From `dataStores.get` when readable, else the id. */
  displayName: string;
  /** Human connector group inferred from the id prefix (SharePoint / OneDrive / Outlook / …). */
  connector: string;
}

export interface DataStoreClientOptions {
  tokens: TokenSource;
  fetchImpl?: FetchLike;
  quotaProject?: string;
  signal?: AbortSignal;
}

/** Infer a friendly connector group from a federated data-store id prefix. */
export function connectorGroup(dataStoreId: string): string {
  const m = /^([a-z0-9]+)-([a-z0-9]+)-fed/i.exec(dataStoreId);
  const source = m?.[2]?.toLowerCase();
  if (!source) return 'Other';
  const label: Record<string, string> = {
    sharepoint: 'SharePoint',
    onedrive: 'OneDrive',
    outlook: 'Outlook',
    salesforce: 'Salesforce',
    jira: 'Jira',
    confluence: 'Confluence',
  };
  return label[source] ?? source.replace(/^\w/, (c) => c.toUpperCase());
}

/** `GET engine` → the `dataStoreIds[]` attached to the app (the discovery path; no `dataStores.list`). */
export async function getEngineDataStoreIds(
  cfg: GeminiClientConfig,
  opts: DataStoreClientOptions,
): Promise<string[]> {
  const raw = await getJson(
    engineUrl(cfg),
    opts.tokens,
    opts.fetchImpl ?? defaultFetch,
    opts.signal,
  );
  return EngineSchema.parse(raw).dataStoreIds ?? [];
}

/** `GET dataStores/{id}` → metadata, or `null` on 404. */
export async function getDataStore(
  cfg: GeminiClientConfig,
  dataStoreId: string,
  opts: DataStoreClientOptions,
): Promise<z.infer<typeof DataStoreSchema> | null> {
  try {
    const raw = await getJson(
      dataStoreUrl(cfg, dataStoreId),
      opts.tokens,
      opts.fetchImpl ?? defaultFetch,
      opts.signal,
    );
    return DataStoreSchema.parse(raw);
  } catch (err) {
    if (err instanceof Error && /\(404\)/.test(err.message)) return null;
    throw err;
  }
}

/**
 * Discover the engine's data stores and enrich each with a display name + connector group — the list a
 * data-store picker renders. `enrich: false` skips the per-store `dataStores.get` (ids + group only).
 */
export async function listEngineDataStores(
  cfg: GeminiClientConfig,
  opts: DataStoreClientOptions & { enrich?: boolean },
): Promise<EngineDataStore[]> {
  const ids = await getEngineDataStoreIds(cfg, opts);
  const enrich = opts.enrich !== false;
  return Promise.all(
    ids.map(async (id): Promise<EngineDataStore> => {
      const resourceName = dataStoreResourceName(cfg.assistant, id);
      let displayName = id;
      if (enrich) {
        try {
          const ds = await getDataStore(cfg, id, opts);
          if (ds?.displayName) displayName = ds.displayName;
        } catch {
          // best-effort enrichment; fall back to the id
        }
      }
      return { id, resourceName, displayName, connector: connectorGroup(id) };
    }),
  );
}

/** Build `streamAssist` `dataStoreSpecs` from selected data store ids (or full resource names). */
export function dataStoreSpecsFromIds(
  cfg: GeminiClientConfig,
  idsOrResourceNames: string[],
): Array<{ dataStore: string }> {
  return idsOrResourceNames.map((v) => ({
    dataStore: v.startsWith('projects/') ? v : dataStoreResourceName(cfg.assistant, v),
  }));
}
