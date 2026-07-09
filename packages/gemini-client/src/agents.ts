import { z } from 'zod';
import {
  assistantResourceName,
  discoveryEngineHost,
  listAvailableAgentViewsUrl,
  type GeminiClientConfig,
} from './config.js';
import { defaultFetch, getJson, postJsonWithHeaders, type FetchLike } from './de-fetch.js';
import type { TokenSource } from './stream-assist.js';

/**
 * Public Discovery Engine **AgentService** lifecycle for skill agents, driven by the
 * signed-in user's Workforce Identity Federation token (client-direct, ADR-0001).
 *
 * Verified live (saib tenant, 2026-07) against `discoveryengine.googleapis.com/v1alpha`:
 * create / get / update / delete and the raw `files:upload` (zip → SKILL.md becomes
 * `instruction`, the rest become `subfiles`) ALL work with the WIF token — **no widget
 * endpoint or widget JWT needed**. This lets the add-in treat skill provisioning as a
 * **boot-time warm-up** (see `ensureSkillAgent`) rather than an out-of-band setup step.
 *
 * Notes from live probing:
 * - `SKILL.md` uploaded in a bundle MUST start with YAML frontmatter (`--- name/description ---`),
 *   else the server returns 400 "Missing YAML frontmatter start delimiter".
 * - Skill agents need NO `:deploy` (that's for managed/app agents only).
 * - Invocation is a separate concern: mount a skill on a turn with the
 *   `mention://?uri=<agentId>` marker in the query (see `stream-assist.ts`) — `skillsSpec`,
 *   `agentsConfig`, and `agentsSpec` do NOT route skill agents on the public endpoint.
 */

export const SkillSubfileSchema = z.object({
  fileName: z.string(),
  mimeType: z.string().optional(),
});
export type SkillSubfile = z.infer<typeof SkillSubfileSchema>;

export const SkillAgentDefinitionSchema = z.object({
  instruction: z.string().optional(),
  gcsUri: z.string().optional(),
  owner: z.string().optional(),
  subfiles: z.array(SkillSubfileSchema).optional(),
});
export type SkillAgentDefinition = z.infer<typeof SkillAgentDefinitionSchema>;

export const AgentSchema = z.object({
  name: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  state: z.string().optional(),
  skillAgentDefinition: SkillAgentDefinitionSchema.optional(),
});
export type Agent = z.infer<typeof AgentSchema>;

export interface AgentServiceOptions {
  /** WIF/OAuth token source (same posture as streamAssist). */
  tokens: TokenSource;
  fetchImpl?: FetchLike;
  /**
   * Billing/quota project for `X-Goog-User-Project`. Only needed when the token does not
   * already carry a quota project (a raw gcloud token does not; a WIF token minted with
   * `WifConfig.userProject` does). Harmless to set either way.
   */
  quotaProject?: string;
  signal?: AbortSignal;
}

function base(cfg: GeminiClientConfig): string {
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/v1alpha/${assistantResourceName(cfg.assistant)}`;
}

function uploadBase(cfg: GeminiClientConfig): string {
  const host = discoveryEngineHost(cfg.assistant.location);
  return `${host}/upload/v1alpha/${assistantResourceName(cfg.assistant)}`;
}

function quotaHeaders(opts: AgentServiceOptions): Record<string, string> {
  return opts.quotaProject ? { 'X-Goog-User-Project': opts.quotaProject } : {};
}

async function sendRaw(
  url: string,
  method: 'PATCH' | 'DELETE',
  opts: AgentServiceOptions,
  body?: unknown,
): Promise<unknown> {
  const fetchImpl = opts.fetchImpl ?? defaultFetch;
  const doSend = async (): Promise<Response> => {
    const token = await opts.tokens.getAccessToken();
    return fetchImpl(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}),
        ...quotaHeaders(opts),
      },
      ...(body !== undefined ? { body: JSON.stringify(body) } : {}),
      ...(opts.signal ? { signal: opts.signal } : {}),
    });
  };
  let res = await doSend();
  if (res.status === 401 && opts.tokens.invalidate) {
    opts.tokens.invalidate();
    res = await doSend();
  }
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(
      `${method} ${url} failed (${res.status})${detail ? `: ${detail.slice(0, 300)}` : ''}`,
    );
  }
  return res.json().catch(() => ({}));
}

/** `GET {assistant}/agents/{id}` → the agent, or `null` if it does not exist (404). */
export async function getAgent(
  cfg: GeminiClientConfig,
  agentId: string,
  opts: AgentServiceOptions,
): Promise<Agent | null> {
  const fetchImpl = opts.fetchImpl ?? defaultFetch;
  try {
    const raw = await getJson(
      `${base(cfg)}/agents/${agentId}`,
      opts.tokens,
      fetchImpl,
      opts.signal,
    );
    return AgentSchema.parse(raw);
  } catch (err) {
    if (err instanceof Error && /\(404\)/.test(err.message)) return null;
    throw err;
  }
}

export const AgentViewSchema = z.object({
  name: z.string().optional(),
  displayName: z.string().optional(),
  description: z.string().optional(),
  state: z.string().optional(),
  agentType: z.string().optional(),
  agentOrigin: z.string().optional(),
  ownerUserPrincipal: z.string().optional(),
});
export type AgentView = z.infer<typeof AgentViewSchema> & { id: string };

const AgentViewsResponseSchema = z.object({
  agentViews: z.array(AgentViewSchema).optional(),
  nextPageToken: z.string().optional(),
});

/**
 * `POST {assistant}:listAvailableAgentViews` — list agents *available to the caller* (owned skills +
 * enabled managed agents). Works with the WIF token (`discoveryengine.agents.listAvailableAgentViews`),
 * unlike `agents.list` (app-gated → 403). This is the taskpane's skill/agent discovery path. Paginates.
 */
export async function listAvailableAgentViews(
  cfg: GeminiClientConfig,
  opts: AgentServiceOptions,
  req?: { pageSize?: number; filter?: string },
): Promise<AgentView[]> {
  const url = listAvailableAgentViewsUrl(cfg);
  const out: AgentView[] = [];
  let pageToken: string | undefined;
  do {
    const raw = await postJsonWithHeaders(
      url,
      {
        pageSize: req?.pageSize ?? 100,
        ...(req?.filter ? { filter: req.filter } : {}),
        ...(pageToken ? { pageToken } : {}),
      },
      opts.tokens,
      opts.fetchImpl ?? defaultFetch,
      quotaHeaders(opts),
      opts.signal,
    );
    const parsed = AgentViewsResponseSchema.parse(raw);
    for (const v of parsed.agentViews ?? []) {
      out.push({ ...v, id: (v.name ?? '').split('/').pop() ?? '' });
    }
    pageToken = parsed.nextPageToken;
  } while (pageToken);
  return out;
}

/** Only the skill agents (`agentType === 'SKILL_AGENT'`) from `listAvailableAgentViews`. */
export async function listSkillAgents(
  cfg: GeminiClientConfig,
  opts: AgentServiceOptions,
): Promise<AgentView[]> {
  const views = await listAvailableAgentViews(cfg, opts);
  return views.filter((v) => v.agentType === 'SKILL_AGENT');
}

/** `POST {assistant}/agents?agentId=<id>` — create a skill agent (inline instruction). */
export async function createSkillAgent(
  cfg: GeminiClientConfig,
  agentId: string,
  agent: { displayName: string; description: string; instruction: string; gcsUri?: string },
  opts: AgentServiceOptions,
): Promise<Agent> {
  const url = `${base(cfg)}/agents?agentId=${encodeURIComponent(agentId)}`;
  const body = {
    displayName: agent.displayName,
    description: agent.description,
    skillAgentDefinition: {
      instruction: agent.instruction,
      ...(agent.gcsUri ? { gcsUri: agent.gcsUri } : {}),
    },
  };
  const raw = await postJsonWithHeaders(
    url,
    body,
    opts.tokens,
    opts.fetchImpl ?? defaultFetch,
    quotaHeaders(opts),
    opts.signal,
  );
  return AgentSchema.parse(raw);
}

/** `PATCH {assistant}/agents/{id}` with a computed updateMask; patches instruction and/or description. */
export async function updateAgent(
  cfg: GeminiClientConfig,
  agentId: string,
  fields: { instruction?: string; description?: string },
  opts: AgentServiceOptions,
): Promise<Agent> {
  const mask: string[] = [];
  const body: { skillAgentDefinition?: { instruction: string }; description?: string } = {};
  if (fields.instruction !== undefined) {
    mask.push('skillAgentDefinition.instruction');
    body.skillAgentDefinition = { instruction: fields.instruction };
  }
  if (fields.description !== undefined) {
    mask.push('description');
    body.description = fields.description;
  }
  const url = `${base(cfg)}/agents/${agentId}?updateMask=${encodeURIComponent(mask.join(','))}`;
  const raw = await sendRaw(url, 'PATCH', opts, body);
  return AgentSchema.parse(raw);
}

/** `DELETE {assistant}/agents/{id}` — removes the agent and its uploaded files. */
export async function deleteAgent(
  cfg: GeminiClientConfig,
  agentId: string,
  opts: AgentServiceOptions,
): Promise<void> {
  await sendRaw(`${base(cfg)}/agents/${agentId}`, 'DELETE', opts);
}

/**
 * Raw bundle upload: `POST /upload/v1alpha/{assistant}/agents/{id}/files:upload?upload_protocol=raw`
 * with `Content-Type: application/zip`. The server unpacks the zip — `SKILL.md` (which MUST have
 * YAML frontmatter) becomes `instruction`, everything else becomes `subfiles`. Re-uploading
 * replaces the fileset (there is no per-subfile delete on the public API).
 */
export async function uploadSkillBundle(
  cfg: GeminiClientConfig,
  agentId: string,
  zipBytes: Uint8Array,
  opts: AgentServiceOptions,
): Promise<void> {
  const url = `${uploadBase(cfg)}/agents/${agentId}/files:upload?upload_protocol=raw`;
  const fetchImpl = opts.fetchImpl ?? defaultFetch;
  const doSend = async (): Promise<Response> => {
    const token = await opts.tokens.getAccessToken();
    return fetchImpl(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/zip',
        ...quotaHeaders(opts),
      },
      body: zipBytes as unknown as BodyInit,
      ...(opts.signal ? { signal: opts.signal } : {}),
    });
  };
  let res = await doSend();
  if (res.status === 401 && opts.tokens.invalidate) {
    opts.tokens.invalidate();
    res = await doSend();
  }
  if (!res.ok) {
    const detail = await res.text().catch(() => '');
    throw new Error(
      `upload ${url} failed (${res.status})${detail ? `: ${detail.slice(0, 300)}` : ''}`,
    );
  }
}

export interface EnsureSkillInput {
  agentId: string;
  displayName: string;
  description: string;
  /**
   * The skill instruction (SKILL.md body). Required to CREATE or UPDATE. When omitted (e.g. the
   * add-in ships only ids + revision from env, not the bundle), the warm-up runs in **detect-only**
   * mode: it reports `missing`/`stale`/`unchanged` without writing.
   */
  instruction?: string;
  /**
   * Revision identifier — the deployer's bundle SHA-256 (e.g. `VITE_GE_*_SHA256`) or a version
   * string. When set, the warm-up stamps it into the agent description as a `[rev:<id>]` marker
   * and, on the next boot, compares the local `revision` against the stored marker — so a changed
   * bundle (instruction OR subfiles) is detected, not just instruction text.
   */
  revision?: string;
  /** Bundle uploaded on create / on drift (zip: SKILL.md→instruction, rest→subfiles). */
  bundleZip?: Uint8Array;
}

export type EnsureSkillResult =
  | { action: 'created'; agent: Agent }
  | { action: 'updated'; agent: Agent }
  | { action: 'unchanged'; agent: Agent }
  /** Detect-only: the skill is absent and no `instruction` was supplied to create it. */
  | { action: 'missing'; agentId: string }
  /** Detect-only: the skill drifted from `revision` and no `instruction` was supplied to update it. */
  | { action: 'stale'; agent: Agent };

/** Trailing `[rev:<id>]` marker used to store the deployed skill revision in the description. */
const REV_MARKER = /\s*\[rev:([^\]]+)\]\s*$/;

function readRevision(description: string | undefined): string | undefined {
  const m = REV_MARKER.exec(description ?? '');
  return m ? m[1] : undefined;
}

function stampRevision(description: string, revision: string): string {
  return `${description.replace(REV_MARKER, '').trimEnd()} [rev:${revision}]`;
}

/**
 * Boot-time warm-up: make the skill agent exist and match `input`, using the signed-in user's WIF
 * token. Idempotent — safe to run on every add-in session start. If the agent is missing it is
 * created (and the bundle uploaded); if the `revision` marker has drifted (or, without a revision,
 * the instruction text changed) it is updated (and the bundle re-uploaded); otherwise nothing is
 * written. Returns what it did so the caller can log/telemeter provisioning.
 */
export async function ensureSkillAgent(
  cfg: GeminiClientConfig,
  input: EnsureSkillInput,
  opts: AgentServiceOptions,
): Promise<EnsureSkillResult> {
  const description = input.revision
    ? stampRevision(input.description, input.revision)
    : input.description;

  const existing = await getAgent(cfg, input.agentId, opts);
  if (!existing) {
    if (input.instruction === undefined) return { action: 'missing', agentId: input.agentId };
    const agent = await createSkillAgent(
      cfg,
      input.agentId,
      { displayName: input.displayName, description, instruction: input.instruction },
      opts,
    );
    if (input.bundleZip) await uploadSkillBundle(cfg, input.agentId, input.bundleZip, opts);
    return { action: 'created', agent };
  }

  const drifted = input.revision
    ? readRevision(existing.description) !== input.revision
    : input.instruction !== undefined &&
      (existing.skillAgentDefinition?.instruction ?? '').trim() !== input.instruction.trim();

  if (drifted) {
    if (input.instruction === undefined) return { action: 'stale', agent: existing };
    const agent = await updateAgent(
      cfg,
      input.agentId,
      { instruction: input.instruction, description },
      opts,
    );
    if (input.bundleZip) await uploadSkillBundle(cfg, input.agentId, input.bundleZip, opts);
    return { action: 'updated', agent };
  }
  return { action: 'unchanged', agent: existing };
}
