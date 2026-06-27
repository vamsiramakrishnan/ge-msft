import { z } from 'zod';
import {
  assistantResourceName,
  type AssistantPath,
  type GeminiSkillMention,
  type WifConfig,
} from '@ge/gemini-client';
import type { ShellConfig } from '../compose.js';
import type { NaaAuthOptions } from '../auth-client.js';
import { ReleaseProfileNameSchema, type ReleaseProfileName } from '@ge/contracts';

export interface RawEnv {
  readonly MODE?: string;
  readonly PROD?: boolean;
  readonly VITE_GCP_PROJECT?: string;
  readonly VITE_GCP_LOCATION?: string;
  readonly VITE_GE_ENGINE?: string;
  readonly VITE_GE_COLLECTION?: string;
  readonly VITE_GE_ASSISTANT?: string;
  readonly VITE_GE_WIDGET_CONFIG_ID?: string;
  readonly VITE_GE_WIDGET_SERVER_TOKEN?: string;
  readonly VITE_GE_MODEL_ID?: string;
  readonly VITE_GE_SKILL_IDS?: string;
  readonly VITE_GE_COMMAND_PLANNER_SKILL?: string;
  readonly VITE_GE_SURFACE_COMMANDER_SKILL?: string;
  readonly VITE_WIF_POOL_ID?: string;
  readonly VITE_WIF_PROVIDER_ID?: string;
  readonly VITE_WIF_SCOPE?: string;
  readonly VITE_WIF_USER_PROJECT?: string;
  readonly VITE_PROXY_URL?: string;
  readonly VITE_ENTRA_TENANT_ID?: string;
  readonly VITE_ENTRA_CLIENT_ID?: string;
  readonly VITE_ENTRA_AUTHORITY?: string;
  readonly VITE_WIF_ID_TOKEN_SCOPES?: string;
  readonly VITE_GRAPH_SCOPES?: string;
  readonly VITE_NOTEBOOK_ID?: string;
  readonly VITE_GE_RELEASE_PROFILE?: string;
}

const ALLOWED_LOCATIONS = [
  'global',
  'us',
  'eu',
  'asia-northeast1',
  'australia-southeast1',
] as const;

const IDENT = /^[a-zA-Z0-9][a-zA-Z0-9._:-]{1,127}$/;
const SKILL_REF = /^[a-zA-Z0-9][a-zA-Z0-9._:/=-]{0,511}$/;
const SKILL_LABEL = /^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/;
const PROJECT = /^([a-z][a-z0-9-]{4,28}[a-z0-9]|\d{6,20})$/;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const WIDGET_TOKEN = /^[a-zA-Z0-9._~+/-]{1,256}$/;

function envKey(...parts: string[]): string {
  return parts.join('_');
}

const FORBIDDEN_BROWSER_KEYS = [
  envKey('ENTRA', 'CLIENT', 'SECRET'),
  envKey('VITE', 'ENTRA', 'CLIENT', 'SECRET'),
  envKey('GOOGLE', 'APPLICATION', 'CREDENTIALS'),
  envKey('GOOGLE', 'CLIENT', 'SECRET'),
  envKey('GCP', 'SERVICE', 'ACCOUNT', 'KEY'),
  envKey('SERVICE', 'ACCOUNT', 'JSON'),
  envKey('GOOGLE', 'PRIVATE', 'KEY'),
  envKey('ACCESS', 'TOKEN'),
  envKey('REFRESH', 'TOKEN'),
] as const;

function required(env: RawEnv, key: keyof RawEnv): string {
  const value = env[key];
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(
      `Missing required environment variable ${key}. Copy .env.example to .env and fill it in.`,
    );
  }
  return value.trim();
}

function isProduction(env: RawEnv): boolean {
  return env.PROD === true || env.MODE === 'production';
}

function ensureNoForbiddenKeys(env: RawEnv | Record<string, unknown>): void {
  const keys = Object.keys(env);
  for (const forbidden of FORBIDDEN_BROWSER_KEYS) {
    if (keys.includes(forbidden)) {
      throw new Error(`${forbidden} must never be present in browser configuration.`);
    }
  }
}

function rejectUnknownProductionKeys(env: RawEnv | Record<string, unknown>): void {
  const allowed = new Set([
    'MODE',
    'PROD',
    'BASE_URL',
    'DEV',
    'SSR',
    ...Object.keys(RawEnvShape.shape),
  ]);
  for (const key of Object.keys(env)) {
    if (key.startsWith('VITE_') && !allowed.has(key)) {
      throw new Error(`Unknown production browser configuration key ${key}.`);
    }
  }
}

const SafeId = z.string().regex(IDENT, 'must be a stable identifier');
const SkillRefs = z.string().regex(SKILL_REF, 'must be a skill id or resource name').array();

const RawEnvShape = z.object({
  VITE_GCP_PROJECT: z.string().regex(PROJECT, 'must be a GCP project id or number'),
  VITE_GCP_LOCATION: z.enum(ALLOWED_LOCATIONS),
  VITE_GE_ENGINE: SafeId,
  VITE_GE_COLLECTION: SafeId.optional(),
  VITE_GE_ASSISTANT: SafeId.optional(),
  VITE_GE_WIDGET_CONFIG_ID: z.string().regex(GUID, 'must be a GE widget config GUID').optional(),
  VITE_GE_WIDGET_SERVER_TOKEN: z.string().regex(WIDGET_TOKEN).optional(),
  VITE_GE_MODEL_ID: SafeId.optional(),
  VITE_GE_SKILL_IDS: SkillRefs.optional(),
  VITE_GE_COMMAND_PLANNER_SKILL: z.string().regex(SKILL_REF).optional(),
  VITE_GE_SURFACE_COMMANDER_SKILL: z.string().regex(SKILL_REF).optional(),
  VITE_WIF_POOL_ID: SafeId,
  VITE_WIF_PROVIDER_ID: SafeId,
  VITE_WIF_SCOPE: z.string().url().optional(),
  VITE_WIF_USER_PROJECT: z.string().regex(PROJECT).optional(),
  VITE_PROXY_URL: z.string().optional(),
  VITE_ENTRA_TENANT_ID: z.string().regex(GUID, 'must be an Entra tenant GUID'),
  VITE_ENTRA_CLIENT_ID: z.string().regex(GUID, 'must be an Entra public-client GUID'),
  VITE_ENTRA_AUTHORITY: z.string().url().optional(),
  VITE_WIF_ID_TOKEN_SCOPES: z.string().optional(),
  VITE_GRAPH_SCOPES: z.string().optional(),
  VITE_NOTEBOOK_ID: SafeId.optional(),
  VITE_GE_RELEASE_PROFILE: ReleaseProfileNameSchema.optional(),
});

function parseEnv(env: RawEnv): z.infer<typeof RawEnvShape> {
  ensureNoForbiddenKeys(env);
  if (isProduction(env)) rejectUnknownProductionKeys(env);
  const raw = {
    VITE_GCP_PROJECT: required(env, 'VITE_GCP_PROJECT'),
    VITE_GCP_LOCATION: required(env, 'VITE_GCP_LOCATION'),
    VITE_GE_ENGINE: required(env, 'VITE_GE_ENGINE'),
    ...(env.VITE_GE_COLLECTION ? { VITE_GE_COLLECTION: env.VITE_GE_COLLECTION.trim() } : {}),
    ...(env.VITE_GE_ASSISTANT ? { VITE_GE_ASSISTANT: env.VITE_GE_ASSISTANT.trim() } : {}),
    ...(env.VITE_GE_WIDGET_CONFIG_ID
      ? { VITE_GE_WIDGET_CONFIG_ID: env.VITE_GE_WIDGET_CONFIG_ID.trim() }
      : {}),
    ...(env.VITE_GE_WIDGET_SERVER_TOKEN
      ? { VITE_GE_WIDGET_SERVER_TOKEN: env.VITE_GE_WIDGET_SERVER_TOKEN.trim() }
      : {}),
    ...(env.VITE_GE_MODEL_ID ? { VITE_GE_MODEL_ID: env.VITE_GE_MODEL_ID.trim() } : {}),
    ...(env.VITE_GE_SKILL_IDS ? { VITE_GE_SKILL_IDS: splitList(env.VITE_GE_SKILL_IDS) } : {}),
    ...(env.VITE_GE_COMMAND_PLANNER_SKILL
      ? { VITE_GE_COMMAND_PLANNER_SKILL: env.VITE_GE_COMMAND_PLANNER_SKILL.trim() }
      : {}),
    ...(env.VITE_GE_SURFACE_COMMANDER_SKILL
      ? { VITE_GE_SURFACE_COMMANDER_SKILL: env.VITE_GE_SURFACE_COMMANDER_SKILL.trim() }
      : {}),
    VITE_WIF_POOL_ID: required(env, 'VITE_WIF_POOL_ID'),
    VITE_WIF_PROVIDER_ID: required(env, 'VITE_WIF_PROVIDER_ID'),
    ...(env.VITE_WIF_SCOPE ? { VITE_WIF_SCOPE: env.VITE_WIF_SCOPE.trim() } : {}),
    ...(env.VITE_WIF_USER_PROJECT
      ? { VITE_WIF_USER_PROJECT: env.VITE_WIF_USER_PROJECT.trim() }
      : {}),
    ...(env.VITE_PROXY_URL
      ? { VITE_PROXY_URL: validateProxyUrl(env.VITE_PROXY_URL, isProduction(env)) }
      : {}),
    VITE_ENTRA_TENANT_ID: required(env, 'VITE_ENTRA_TENANT_ID'),
    VITE_ENTRA_CLIENT_ID: required(env, 'VITE_ENTRA_CLIENT_ID'),
    ...(env.VITE_ENTRA_AUTHORITY
      ? {
          VITE_ENTRA_AUTHORITY: validateAuthority(
            env.VITE_ENTRA_AUTHORITY,
            env.VITE_ENTRA_TENANT_ID,
            isProduction(env),
          ),
        }
      : {}),
    ...(env.VITE_WIF_ID_TOKEN_SCOPES
      ? { VITE_WIF_ID_TOKEN_SCOPES: env.VITE_WIF_ID_TOKEN_SCOPES }
      : {}),
    ...(env.VITE_GRAPH_SCOPES ? { VITE_GRAPH_SCOPES: env.VITE_GRAPH_SCOPES } : {}),
    ...(env.VITE_NOTEBOOK_ID ? { VITE_NOTEBOOK_ID: env.VITE_NOTEBOOK_ID.trim() } : {}),
    ...(env.VITE_GE_RELEASE_PROFILE
      ? { VITE_GE_RELEASE_PROFILE: env.VITE_GE_RELEASE_PROFILE }
      : {}),
  };
  return RawEnvShape.parse(raw);
}

function validateHttpsUrl(raw: string, name: string, prod: boolean): string {
  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    throw new Error(`${name} must be a valid URL.`);
  }
  if (parsed.username || parsed.password) throw new Error(`${name} must not contain credentials.`);
  const local = parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';
  if (prod && local) throw new Error(`${name} must not use localhost in production.`);
  if (prod && parsed.protocol !== 'https:')
    throw new Error(`${name} must use https in production.`);
  if (!prod && parsed.protocol !== 'https:' && !(parsed.protocol === 'http:' && local)) {
    throw new Error(`${name} must use https except for localhost development.`);
  }
  if (!['https:', 'http:'].includes(parsed.protocol))
    throw new Error(`${name} uses an unsafe scheme.`);
  return parsed.toString().replace(/\/$/, '');
}

function validateProxyUrl(raw: string, prod: boolean): string {
  return validateHttpsUrl(raw, 'VITE_PROXY_URL', prod);
}

function validateAuthority(raw: string, tenantId: string | undefined, prod: boolean): string {
  const value = validateHttpsUrl(raw, 'VITE_ENTRA_AUTHORITY', prod);
  const parsed = new URL(value);
  if (prod && parsed.hostname !== 'login.microsoftonline.com') {
    throw new Error('VITE_ENTRA_AUTHORITY must use login.microsoftonline.com in production.');
  }
  if (tenantId && !parsed.pathname.toLowerCase().includes(tenantId.toLowerCase())) {
    throw new Error('VITE_ENTRA_AUTHORITY must include VITE_ENTRA_TENANT_ID.');
  }
  return value;
}

function splitScopes(value: string | undefined, fallback: string[]): string[] {
  if (!value) return fallback;
  const scopes = splitList(value);
  return scopes.length ? scopes : fallback;
}

function splitList(value: string): string[] {
  return value
    .split(/[, ]+/)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

const OIDC_ONLY_SCOPES = new Set(['offline_access', 'openid', 'profile', 'email']);

function requireNaaResourceScope(scopes: string[], name: string): string[] {
  if (scopes.every((scope) => OIDC_ONLY_SCOPES.has(scope))) {
    throw new Error(
      `${name} must include at least one non-OIDC scope such as User.Read or an app resource scope.`,
    );
  }
  return scopes;
}

export function assistantFromEnv(env: RawEnv): AssistantPath {
  const parsed = parseEnv(env);
  return {
    project: parsed.VITE_GCP_PROJECT,
    location: parsed.VITE_GCP_LOCATION,
    engine: parsed.VITE_GE_ENGINE,
    ...(parsed.VITE_GE_COLLECTION ? { collection: parsed.VITE_GE_COLLECTION } : {}),
    ...(parsed.VITE_GE_ASSISTANT ? { assistant: parsed.VITE_GE_ASSISTANT } : {}),
  };
}

export function wifFromEnv(env: RawEnv): WifConfig {
  const parsed = parseEnv(env);
  return {
    poolId: parsed.VITE_WIF_POOL_ID,
    providerId: parsed.VITE_WIF_PROVIDER_ID,
    ...(parsed.VITE_WIF_SCOPE ? { scope: parsed.VITE_WIF_SCOPE } : {}),
    ...(parsed.VITE_WIF_USER_PROJECT ? { userProject: parsed.VITE_WIF_USER_PROJECT } : {}),
  };
}

export function shellConfigFromEnv(env: RawEnv): ShellConfig {
  const parsed = parseEnv(env);
  return {
    assistant: assistantFromEnv(env),
    ...(widgetFromEnv(env) ? { widget: widgetFromEnv(env) } : {}),
    wif: wifFromEnv(env),
    ...(parsed.VITE_GE_MODEL_ID ? { modelId: parsed.VITE_GE_MODEL_ID } : {}),
    ...(parsed.VITE_GE_SKILL_IDS?.length ? { skills: skillsFromEnv(env) } : {}),
    ...(parsed.VITE_GE_SKILL_IDS?.length ? { skillMentions: skillMentionsFromEnv(env) } : {}),
    ...(parsed.VITE_GE_COMMAND_PLANNER_SKILL ? { plannerSkills: plannerSkillsFromEnv(env) } : {}),
    ...(parsed.VITE_GE_COMMAND_PLANNER_SKILL
      ? { plannerSkillMentions: plannerSkillMentionsFromEnv(env) }
      : {}),
    ...(parsed.VITE_GE_SURFACE_COMMANDER_SKILL ? { commandSkills: commandSkillsFromEnv(env) } : {}),
    ...(parsed.VITE_GE_SURFACE_COMMANDER_SKILL
      ? { commandSkillMentions: commandSkillMentionsFromEnv(env) }
      : {}),
    ...(parsed.VITE_PROXY_URL ? { proxyUrl: parsed.VITE_PROXY_URL } : {}),
    ...(parsed.VITE_GE_RELEASE_PROFILE ? { releaseProfile: parsed.VITE_GE_RELEASE_PROFILE } : {}),
  };
}

export function widgetFromEnv(env: RawEnv): ShellConfig['widget'] | undefined {
  const parsed = parseEnv(env);
  if (!parsed.VITE_GE_WIDGET_CONFIG_ID) return undefined;
  return {
    configId: parsed.VITE_GE_WIDGET_CONFIG_ID,
    ...(parsed.VITE_GE_WIDGET_SERVER_TOKEN
      ? { serverToken: parsed.VITE_GE_WIDGET_SERVER_TOKEN }
      : {}),
  };
}

export function skillsFromEnv(env: RawEnv): string[] {
  return skillEntriesFromEnv(env).map((skill) => skill.resource);
}

export function skillMentionsFromEnv(env: RawEnv): GeminiSkillMention[] {
  return skillEntriesFromEnv(env)
    .filter(
      (skill): skill is { resource: string; mention: GeminiSkillMention } => skill.mention != null,
    )
    .map((skill) => skill.mention);
}

export function plannerSkillsFromEnv(env: RawEnv): string[] {
  return routeSkillFromEnv(env, 'VITE_GE_COMMAND_PLANNER_SKILL').map((skill) => skill.resource);
}

export function plannerSkillMentionsFromEnv(env: RawEnv): GeminiSkillMention[] {
  return routeSkillFromEnv(env, 'VITE_GE_COMMAND_PLANNER_SKILL')
    .map((skill) => skill.mention)
    .filter((mention): mention is GeminiSkillMention => mention !== undefined);
}

export function commandSkillsFromEnv(env: RawEnv): string[] {
  return routeSkillFromEnv(env, 'VITE_GE_SURFACE_COMMANDER_SKILL').map((skill) => skill.resource);
}

export function commandSkillMentionsFromEnv(env: RawEnv): GeminiSkillMention[] {
  return routeSkillFromEnv(env, 'VITE_GE_SURFACE_COMMANDER_SKILL')
    .map((skill) => skill.mention)
    .filter((mention): mention is GeminiSkillMention => mention !== undefined);
}

function skillEntriesFromEnv(env: RawEnv): { resource: string; mention?: GeminiSkillMention }[] {
  const parsed = parseEnv(env);
  const skills = parsed.VITE_GE_SKILL_IDS ?? [];
  if (skills.length === 0) return [];
  const assistant = assistantResourceName(assistantFromEnv(env));
  return skills.map((entry) => parseSkillEntry(entry, assistant));
}

function routeSkillFromEnv(
  env: RawEnv,
  key: 'VITE_GE_COMMAND_PLANNER_SKILL' | 'VITE_GE_SURFACE_COMMANDER_SKILL',
): { resource: string; mention?: GeminiSkillMention }[] {
  const parsed = parseEnv(env);
  const entry = parsed[key];
  if (!entry) return [];
  return [parseSkillEntry(entry, assistantResourceName(assistantFromEnv(env)))];
}

function parseSkillEntry(
  entry: string,
  assistant: string,
): { resource: string; mention?: GeminiSkillMention } {
  const delimiter = entry.indexOf('=');
  const label = delimiter >= 0 ? entry.slice(0, delimiter).trim() : undefined;
  const skill = delimiter >= 0 ? entry.slice(delimiter + 1).trim() : entry.trim();
  if (!skill) throw new Error('VITE_GE_SKILL_IDS contains an empty skill resource.');
  const resource = skill.startsWith('projects/') ? skill : `${assistant}/agents/${skill}`;
  if (!label) return { resource };
  if (!SKILL_LABEL.test(label)) {
    throw new Error(`Invalid Gemini Enterprise skill mention label ${label}.`);
  }
  const uri = agentIdFromResource(resource);
  return { resource, mention: { label, uri } };
}

function agentIdFromResource(resource: string): string {
  const marker = '/agents/';
  const at = resource.lastIndexOf(marker);
  if (at < 0) throw new Error(`Gemini Enterprise skill resource must contain ${marker}.`);
  const id = resource.slice(at + marker.length);
  if (!id || id.includes('/')) {
    throw new Error('Gemini Enterprise skill resource has an invalid agent id.');
  }
  return id;
}

export function authOptionsFromEnv(env: RawEnv): NaaAuthOptions {
  const parsed = parseEnv(env);
  const idTokenScopes = requireNaaResourceScope(
    splitScopes(parsed.VITE_WIF_ID_TOKEN_SCOPES, [`${parsed.VITE_ENTRA_CLIENT_ID}/.default`]),
    'VITE_WIF_ID_TOKEN_SCOPES',
  );
  const graphScopes = splitScopes(parsed.VITE_GRAPH_SCOPES, ['User.Read']);
  return { idTokenScopes, graphScopes };
}

export interface MsalConfigLike {
  clientId: string;
  authority?: string;
  redirectUri?: string;
  popupBridgeTimeoutMs?: number;
  iframeBridgeTimeoutMs?: number;
  forceStandardPopupBridge?: boolean;
}

export function msalConfigFromEnv(env: RawEnv): MsalConfigLike {
  const parsed = parseEnv(env);
  return {
    clientId: parsed.VITE_ENTRA_CLIENT_ID,
    ...(parsed.VITE_ENTRA_AUTHORITY ? { authority: parsed.VITE_ENTRA_AUTHORITY } : {}),
  };
}

export function notebookIdFromEnv(env: RawEnv): string | undefined {
  return parseEnv(env).VITE_NOTEBOOK_ID;
}

export function releaseProfileFromEnv(env: RawEnv): ReleaseProfileName | undefined {
  return parseEnv(env).VITE_GE_RELEASE_PROFILE;
}
