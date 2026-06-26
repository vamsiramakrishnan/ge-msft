import { z } from 'zod';
import {
  assistantAgentsUrl,
  dataStoresUrl,
  lookupWidgetConfigUrl,
  widgetListAvailableAgentViewsUrl,
  type GeminiClientConfig,
  type GeminiSkillMention,
} from './config.js';
import { defaultFetch, getJson, postJsonWithHeaders, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';
import type { RetryOptions } from './retry.js';

export interface GeminiCatalogSkill {
  name: string;
  id: string;
  label: string;
  description?: string;
  state?: string;
  mention: GeminiSkillMention;
  suggestedRoute?: 'planner' | 'command';
}

export interface GeminiCatalogDataStore {
  name: string;
  id: string;
  label: string;
  type?: string;
  suggested?: boolean;
}

export interface GeminiCatalog {
  skills: GeminiCatalogSkill[];
  dataStores: GeminiCatalogDataStore[];
  warnings?: string[];
}

export interface GeminiCatalogSelection {
  defaultSkills?: GeminiCatalogSkill[];
  plannerSkill?: GeminiCatalogSkill;
  commandSkill?: GeminiCatalogSkill;
  dataStores?: GeminiCatalogDataStore[];
}

const AgentSchema = z
  .object({
    name: z.string(),
    displayName: z.string().optional(),
    description: z.string().optional(),
    state: z.string().optional(),
  })
  .passthrough();

const AgentListSchema = z
  .object({
    agents: z.array(AgentSchema).optional(),
    skills: z.array(AgentSchema).optional(),
  })
  .passthrough();

const WidgetAgentViewSchema = z
  .object({
    name: z.string().optional(),
    displayName: z.string().optional(),
    description: z.string().optional(),
    agentId: z.union([z.string(), z.number()]).optional(),
    uri: z.union([z.string(), z.number()]).optional(),
    agent: AgentSchema.optional(),
    agentView: z.unknown().optional(),
  })
  .passthrough();

const WidgetAgentViewsResponseSchema = z
  .object({
    agentViews: z.array(WidgetAgentViewSchema).optional(),
    availableAgentViews: z.array(WidgetAgentViewSchema).optional(),
    agents: z.array(WidgetAgentViewSchema).optional(),
    nextPageToken: z.string().optional(),
  })
  .passthrough();

const DataStoreSchema = z
  .object({
    name: z.string(),
    displayName: z.string().optional(),
    type: z.string().optional(),
  })
  .passthrough();

const DataStoreListSchema = z
  .object({
    dataStores: z.array(DataStoreSchema).optional(),
  })
  .passthrough();

const DATA_STORE_RESOURCE =
  /^(?:projects\/[^/]+\/locations\/[^/]+\/)?collections\/[^/]+\/dataStores\/[^/]+$/;
const DATA_STORE_ID = /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/;

export class DiscoveryCatalogClient {
  constructor(
    private readonly tokens: TokenSource,
    private readonly config: GeminiClientConfig,
    private readonly fetchImpl: FetchLike = defaultFetch,
    private readonly retryOpts: RetryOptions = {},
  ) {}

  async listCatalog(signal?: AbortSignal): Promise<GeminiCatalog> {
    const configured = configuredCatalog(this.config);
    const [skillsResult, dataStoresResult] = await Promise.allSettled([
      this.listSkills(signal),
      this.listDataStores(signal),
    ]);
    const warnings: string[] = [];
    const skills = skillsResult.status === 'fulfilled' ? skillsResult.value : configured.skills;
    const dataStores =
      dataStoresResult.status === 'fulfilled' ? dataStoresResult.value : configured.dataStores;

    if (skillsResult.status === 'rejected') {
      warnings.push(
        catalogWarning(
          'skills',
          this.config.widget?.configId
            ? 'content-discoveryengine.widgetListAvailableAgentViews'
            : 'discoveryengine.agents.list',
          skillsResult.reason,
          configured.skills.length,
        ),
      );
    }
    if (dataStoresResult.status === 'rejected') {
      warnings.push(
        catalogWarning(
          'connectors',
          this.config.widget?.configId
            ? 'content-discoveryengine.lookupWidgetConfig or discoveryengine.dataStores.list'
            : 'discoveryengine.dataStores.list',
          dataStoresResult.reason,
          configured.dataStores.length,
        ),
      );
    }

    return {
      skills,
      dataStores,
      ...(warnings.length ? { warnings } : {}),
    };
  }

  async listSkills(signal?: AbortSignal): Promise<GeminiCatalogSkill[]> {
    if (this.config.widget?.configId) {
      return this.listWidgetSkills(signal);
    }
    return this.listAdminSkills(signal);
  }

  async listAdminSkills(signal?: AbortSignal): Promise<GeminiCatalogSkill[]> {
    const raw = await getJson(
      assistantAgentsUrl(this.config),
      this.tokens,
      this.fetchImpl,
      signal,
      this.retryOpts,
    );
    const parsed = AgentListSchema.parse(raw);
    const agents = parsed.agents ?? parsed.skills ?? [];
    return agents.map((agent) => normalizeSkill(agent.name, agent));
  }

  async listWidgetSkills(signal?: AbortSignal): Promise<GeminiCatalogSkill[]> {
    const origins = ['USER', 'GOOGLE'] as const;
    const views = (
      await Promise.all(origins.map((origin) => this.listWidgetSkillOrigin(origin, signal)))
    ).flat();
    const deduped = new Map<string, GeminiCatalogSkill>();
    for (const skill of views) deduped.set(skill.name, skill);
    return [...deduped.values()];
  }

  private async listWidgetSkillOrigin(
    agentOrigin: 'USER' | 'GOOGLE',
    signal?: AbortSignal,
  ): Promise<GeminiCatalogSkill[]> {
    const widget = this.config.widget;
    if (!widget?.configId) return [];
    const raw = await postJsonWithHeaders(
      widgetListAvailableAgentViewsUrl(this.config),
      {
        configId: widget.configId,
        additionalParams: { token: '-', origin: 'ORIGIN_UNSPECIFIED' },
        listAvailableAgentViewsRequest: {
          pageSize: 200,
          filter: 'agent_type = SKILL_AGENT',
          agentOrigin,
        },
      },
      this.tokens,
      this.fetchImpl,
      widget.serverToken ? { 'x-server-token': widget.serverToken } : {},
      signal,
      this.retryOpts,
    );
    const parsed = WidgetAgentViewsResponseSchema.parse(raw);
    const entries = parsed.agentViews ?? parsed.availableAgentViews ?? parsed.agents ?? [];
    return entries.map((entry) => normalizeWidgetSkill(entry));
  }

  async listDataStores(signal?: AbortSignal): Promise<GeminiCatalogDataStore[]> {
    let widgetError: unknown;
    if (this.config.widget?.configId) {
      try {
        const widgetStores = await this.listWidgetDataStores(signal);
        if (widgetStores.length > 0) return widgetStores;
      } catch (err) {
        widgetError = err;
      }
    }
    try {
      return await this.listAdminDataStores(signal);
    } catch (err) {
      if (widgetError) {
        throw new Error(
          `Widget config lookup failed: ${errorText(widgetError)}; dataStores.list failed: ${errorText(err)}`,
        );
      }
      throw err;
    }
  }

  async listWidgetDataStores(signal?: AbortSignal): Promise<GeminiCatalogDataStore[]> {
    const widget = this.config.widget;
    if (!widget?.configId) return [];
    const raw = await postJsonWithHeaders(
      lookupWidgetConfigUrl(this.config),
      { widgetConfigId: widget.configId },
      this.tokens,
      this.fetchImpl,
      widget.serverToken ? { 'x-server-token': widget.serverToken } : {},
      signal,
      this.retryOpts,
    );
    return extractWidgetDataStores(raw, this.config);
  }

  async listAdminDataStores(signal?: AbortSignal): Promise<GeminiCatalogDataStore[]> {
    const raw = await getJson(
      dataStoresUrl(this.config),
      this.tokens,
      this.fetchImpl,
      signal,
      this.retryOpts,
    );
    const parsed = DataStoreListSchema.parse(raw);
    return (parsed.dataStores ?? []).map((store) => normalizeDataStore(store.name, store));
  }
}

export function configuredCatalog(config: GeminiClientConfig): GeminiCatalog {
  const skills = new Map<string, GeminiCatalogSkill>();
  addConfiguredSkills(skills, config.skills ?? [], config.skillMentions ?? []);
  addConfiguredSkills(
    skills,
    config.plannerSkills ?? [],
    config.plannerSkillMentions ?? [],
    'planner',
  );
  addConfiguredSkills(
    skills,
    config.commandSkills ?? [],
    config.commandSkillMentions ?? [],
    'command',
  );
  return {
    skills: [...skills.values()],
    dataStores: (config.dataStores ?? []).map((name) => normalizeDataStore(name, {})),
  };
}

export function applyCatalogSelection(
  selection: GeminiCatalogSelection,
): Partial<
  Pick<
    GeminiClientConfig,
    | 'skills'
    | 'skillMentions'
    | 'plannerSkills'
    | 'plannerSkillMentions'
    | 'commandSkills'
    | 'commandSkillMentions'
    | 'dataStores'
  >
> {
  const defaultSkills = selection.defaultSkills ?? [];
  const planner = selection.plannerSkill ? [selection.plannerSkill] : [];
  const command = selection.commandSkill ? [selection.commandSkill] : [];
  return {
    skills: defaultSkills.map((s) => s.name),
    skillMentions: defaultSkills.map((s) => s.mention),
    plannerSkills: planner.map((s) => s.name),
    plannerSkillMentions: planner.map((s) => s.mention),
    commandSkills: command.map((s) => s.name),
    commandSkillMentions: command.map((s) => s.mention),
    dataStores: (selection.dataStores ?? []).map((d) => d.name),
  };
}

export function defaultCatalogSelection(catalog: GeminiCatalog): GeminiCatalogSelection {
  return {
    plannerSkill: catalog.skills.find((s) => s.suggestedRoute === 'planner'),
    commandSkill: catalog.skills.find((s) => s.suggestedRoute === 'command'),
    dataStores: catalog.dataStores.filter((store) => store.suggested),
  };
}

function normalizeSkill(
  name: string,
  agent: { displayName?: string; description?: string; state?: string },
): GeminiCatalogSkill {
  const id = tail(name, '/agents/');
  const label = agent.displayName?.trim() || id;
  const route = suggestedRoute(label, id);
  return {
    name,
    id,
    label,
    ...(agent.description ? { description: agent.description } : {}),
    ...(agent.state ? { state: agent.state } : {}),
    mention: { label, uri: id },
    ...(route ? { suggestedRoute: route } : {}),
  };
}

function normalizeWidgetSkill(view: z.infer<typeof WidgetAgentViewSchema>): GeminiCatalogSkill {
  const nested = view.agent ?? extractNestedAgent(view.agentView);
  const name =
    nested?.name ??
    view.name ??
    (view.agentId || view.uri
      ? `${assistantNameFromUnknownView(view)}/agents/${String(view.agentId ?? view.uri)}`
      : undefined);
  if (!name) {
    throw new Error('Widget agent view did not include a skill name or id.');
  }
  const id = String(view.uri ?? view.agentId ?? tail(name, '/agents/'));
  const label = view.displayName?.trim() || nested?.displayName?.trim() || id;
  const route = suggestedRoute(label, id);
  return {
    name,
    id,
    label,
    ...((view.description ?? nested?.description)
      ? { description: view.description ?? nested?.description }
      : {}),
    ...(nested?.state ? { state: nested.state } : {}),
    mention: { label, uri: id },
    ...(route ? { suggestedRoute: route } : {}),
  };
}

function extractNestedAgent(value: unknown): z.infer<typeof AgentSchema> | undefined {
  const parsed = AgentSchema.safeParse(value);
  return parsed.success ? parsed.data : undefined;
}

function assistantNameFromUnknownView(view: z.infer<typeof WidgetAgentViewSchema>): string {
  const maybeName = typeof view.name === 'string' ? view.name : undefined;
  const marker = '/agents/';
  if (maybeName?.includes(marker)) return maybeName.slice(0, maybeName.indexOf(marker));
  throw new Error('Widget agent view with only an id must also include an agent resource name.');
}

function normalizeDataStore(
  name: string,
  store: { displayName?: string; type?: string },
): GeminiCatalogDataStore {
  const id = tail(name, '/dataStores/');
  const label = store.displayName?.trim() || id;
  return {
    name,
    id,
    label,
    ...(store.type ? { type: store.type } : {}),
    ...(suggestedDataStore(label, id) ? { suggested: true } : {}),
  };
}

function extractWidgetDataStores(
  raw: unknown,
  config: GeminiClientConfig,
): GeminiCatalogDataStore[] {
  const stores = new Map<string, { name: string; displayName?: string; type?: string }>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (!isRecord(value)) return;
    const name = dataStoreNameFromWidgetObject(value, config);
    if (name && !stores.has(name)) {
      stores.set(name, {
        name,
        ...dataStoreMetadataFromWidgetObject(value),
      });
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(raw);
  return [...stores.values()].map((store) => normalizeDataStore(store.name, store));
}

function dataStoreNameFromWidgetObject(
  value: Record<string, unknown>,
  config: GeminiClientConfig,
): string | undefined {
  const directId = stringValue(value.dataStoreId);
  if (directId && DATA_STORE_ID.test(directId)) {
    const collection = config.assistant.collection ?? 'default_collection';
    return `projects/${config.assistant.project}/locations/${config.assistant.location}/collections/${collection}/dataStores/${directId}`;
  }
  const candidates = [
    value.dataStore,
    value.dataStoreName,
    value.dataStoreId,
    value.name,
    value.resourceName,
  ];
  for (const candidate of candidates) {
    if (typeof candidate !== 'string') continue;
    const normalized = normalizeWidgetDataStoreName(candidate, config);
    if (normalized) return normalized;
  }
  return undefined;
}

function normalizeWidgetDataStoreName(
  value: string,
  config: GeminiClientConfig,
): string | undefined {
  const trimmed = value.trim();
  if (!DATA_STORE_RESOURCE.test(trimmed)) return undefined;
  if (trimmed.startsWith('projects/')) return trimmed;
  return `projects/${config.assistant.project}/locations/${config.assistant.location}/${trimmed}`;
}

function dataStoreMetadataFromWidgetObject(value: Record<string, unknown>): {
  displayName?: string;
  type?: string;
} {
  const displayName =
    stringValue(value.displayName) ??
    stringValue(value.display_name) ??
    stringValue(value.title) ??
    stringValue(value.label);
  const type = stringValue(value.type) ?? stringValue(value.connectorType);
  return {
    ...(displayName ? { displayName } : {}),
    ...(type ? { type } : {}),
  };
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function addConfiguredSkills(
  out: Map<string, GeminiCatalogSkill>,
  resources: readonly string[],
  mentions: readonly GeminiSkillMention[],
  route?: GeminiCatalogSkill['suggestedRoute'],
): void {
  for (const resource of resources) {
    const mention = mentions.find((m) => m.uri === tail(resource, '/agents/'));
    const skill = normalizeSkill(resource, {
      ...(mention ? { displayName: mention.label } : {}),
    });
    out.set(resource, route ? { ...skill, suggestedRoute: route } : skill);
  }
}

function catalogWarning(
  noun: string,
  permission: string,
  reason: unknown,
  fallbackCount: number,
): string {
  const text = reason instanceof Error ? reason.message : String(reason);
  const permissionDenied = /\b403\b|PERMISSION_DENIED|permission/i.test(text);
  const fallback =
    fallbackCount > 0
      ? ` Using ${fallbackCount} configured fallback ${noun}.`
      : ' No configured fallback entries were available.';
  if (permissionDenied) {
    return `Could not list GE ${noun}: the signed-in user needs ${permission}.${fallback}`;
  }
  return `Could not list GE ${noun}.${fallback}`;
}

function suggestedRoute(label: string, id: string): 'planner' | 'command' | undefined {
  const text = `${label} ${id}`.toLowerCase();
  if (text.includes('surface-commander') || text.includes('commander')) return 'command';
  if (text.includes('command-planner') || /\bplanner\b/.test(text)) return 'planner';
  return undefined;
}

function suggestedDataStore(label: string, id: string): boolean {
  const text = `${label} ${id}`.toLowerCase();
  return (
    text.includes('onedrive') ||
    text.includes('sharepoint') ||
    text.includes('outlook') ||
    text.includes('mail') ||
    text.includes('calendar') ||
    text.includes('contact')
  );
}

function tail(value: string, marker: string): string {
  const at = value.lastIndexOf(marker);
  return at >= 0 ? value.slice(at + marker.length) : (value.split('/').pop() ?? value);
}
