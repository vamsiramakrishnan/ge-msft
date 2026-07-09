import { describe, it, expect } from 'vitest';
import {
  assistantFromEnv,
  wifFromEnv,
  shellConfigFromEnv,
  authOptionsFromEnv,
  msalConfigFromEnv,
  notebookIdFromEnv,
  releaseProfileFromEnv,
  commandSkillMentionsFromEnv,
  commandSkillsFromEnv,
  plannerSkillMentionsFromEnv,
  plannerSkillsFromEnv,
  skillMentionsFromEnv,
  skillsFromEnv,
  warmUpSkillsFromEnv,
  widgetFromEnv,
  type RawEnv,
} from './config.js';

const SHA_A = 'a'.repeat(64);
const SHA_B = 'b'.repeat(64);
const SHA_C = 'c'.repeat(64);
const SHA_D = 'd'.repeat(64);
const SHA_E = 'e'.repeat(64);
const SHA_F = 'f'.repeat(64);

const full: RawEnv = {
  MODE: 'production',
  PROD: true,
  VITE_GCP_PROJECT: 'proj-12345',
  VITE_GCP_LOCATION: 'eu',
  VITE_GE_ENGINE: 'engine-1',
  VITE_GE_COLLECTION: 'col-1',
  VITE_GE_ASSISTANT: 'assist-1',
  VITE_GE_WIDGET_CONFIG_ID: '33333333-3333-4333-8333-333333333333',
  VITE_GE_WIDGET_SERVER_TOKEN: 'test-widget-server-token',
  VITE_GE_MODEL_ID: 'gemini-x',
  VITE_GE_COMMAND_PLANNER_SKILL:
    'm365-command-planner=projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/17573173582293271726',
  VITE_GE_COMMAND_PLANNER_SKILL_VERSION: '1.1',
  VITE_GE_COMMAND_PLANNER_SKILL_SOURCE_SHA256: SHA_A,
  VITE_GE_COMMAND_PLANNER_SKILL_SHA256: SHA_B,
  VITE_GE_SURFACE_COMMANDER_SKILL:
    'm365-surface-commander=projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129',
  VITE_GE_SURFACE_COMMANDER_SKILL_VERSION: '1.1',
  VITE_GE_SURFACE_COMMANDER_SKILL_SOURCE_SHA256: SHA_C,
  VITE_GE_SURFACE_COMMANDER_SKILL_SHA256: SHA_D,
  VITE_GE_SKILL_SOURCE_BUNDLE_SET_SHA256: SHA_E,
  VITE_GE_SKILL_UPLOAD_BUNDLE_SET_SHA256: SHA_F,
  VITE_WIF_POOL_ID: 'pool-1',
  VITE_WIF_PROVIDER_ID: 'prov-1',
  VITE_WIF_SCOPE: 'https://www.googleapis.com/auth/cloud-platform',
  VITE_WIF_USER_PROJECT: 'billing-12345',
  VITE_PROXY_URL: 'https://proxy.contoso.example',
  VITE_ENTRA_TENANT_ID: '11111111-1111-4111-8111-111111111111',
  VITE_ENTRA_CLIENT_ID: '22222222-2222-4222-8222-222222222222',
  VITE_ENTRA_AUTHORITY: 'https://login.microsoftonline.com/11111111-1111-4111-8111-111111111111',
  VITE_WIF_ID_TOKEN_SCOPES: 'openid profile User.Read',
  VITE_GRAPH_SCOPES: 'User.Read Files.Read.Selected',
  VITE_NOTEBOOK_ID: 'nb-1',
  VITE_GE_RELEASE_PROFILE: 'internal-alpha-word-excel',
};

describe('warmUpSkillsFromEnv', () => {
  it('builds detect-only warm-up inputs (id + display + revision, no instruction) for both skills', () => {
    const warmUp = warmUpSkillsFromEnv(full);
    expect(warmUp).toEqual([
      {
        agentId: '17573173582293271726',
        displayName: 'm365-command-planner',
        description: 'm365-command-planner — Gemini Enterprise skill for the M365 add-in',
        revision: SHA_B,
      },
      {
        agentId: '7404511736383961129',
        displayName: 'm365-surface-commander',
        description: 'm365-surface-commander — Gemini Enterprise skill for the M365 add-in',
        revision: SHA_D,
      },
    ]);
    // detect-only: no instruction/bundle shipped from env
    expect(warmUp.every((s) => s.instruction === undefined && s.bundleZip === undefined)).toBe(
      true,
    );
  });

  it('is empty when no skills are configured', () => {
    const env: RawEnv = { ...full };
    delete (env as Record<string, unknown>).VITE_GE_COMMAND_PLANNER_SKILL;
    delete (env as Record<string, unknown>).VITE_GE_SURFACE_COMMANDER_SKILL;
    expect(warmUpSkillsFromEnv(env)).toEqual([]);
  });
});

describe('config from env', () => {
  it('builds the assistant path with optional fields', () => {
    expect(assistantFromEnv(full)).toEqual({
      project: 'proj-12345',
      location: 'eu',
      engine: 'engine-1',
      collection: 'col-1',
      assistant: 'assist-1',
    });
  });

  it('omits optional assistant fields when unset', () => {
    const { VITE_GE_COLLECTION: _c, VITE_GE_ASSISTANT: _a, ...rest } = full;
    const path = assistantFromEnv(rest);
    expect(path).not.toHaveProperty('collection');
    expect(path).not.toHaveProperty('assistant');
  });

  it('builds the WIF config', () => {
    expect(wifFromEnv(full)).toEqual({
      poolId: 'pool-1',
      providerId: 'prov-1',
      scope: 'https://www.googleapis.com/auth/cloud-platform',
      userProject: 'billing-12345',
    });
  });

  it('assembles the full shell config with release profile and no secrets', () => {
    const cfg = shellConfigFromEnv(full);
    expect(cfg.assistant.project).toBe('proj-12345');
    expect(cfg.widget).toEqual({
      configId: '33333333-3333-4333-8333-333333333333',
      serverToken: 'test-widget-server-token',
    });
    expect(cfg.modelId).toBe('gemini-x');
    expect(cfg.plannerSkills).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/17573173582293271726',
    ]);
    expect(cfg.plannerSkillMentions).toEqual([
      { label: 'm365-command-planner', uri: '17573173582293271726' },
    ]);
    expect(cfg.commandSkills).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129',
    ]);
    expect(cfg.commandSkillMentions).toEqual([
      { label: 'm365-surface-commander', uri: '7404511736383961129' },
    ]);
    expect(cfg.proxyUrl).toBe('https://proxy.contoso.example');
    expect(cfg.releaseProfile).toBe('internal-alpha-word-excel');
    expect(JSON.stringify(cfg)).not.toMatch(/secret|private_key|refresh_token/i);
  });

  it('parses space/comma scopes; falls back to defaults', () => {
    expect(authOptionsFromEnv(full).idTokenScopes).toEqual(['openid', 'profile', 'User.Read']);
    expect(authOptionsFromEnv(full).graphScopes).toEqual(['User.Read', 'Files.Read.Selected']);
    const noScopes = { ...full, VITE_GRAPH_SCOPES: undefined };
    expect(authOptionsFromEnv(noScopes).graphScopes).toEqual(['User.Read']);
  });

  it('rejects NAA WIF token scopes that contain only OIDC scopes', () => {
    expect(() =>
      authOptionsFromEnv({ ...full, VITE_WIF_ID_TOKEN_SCOPES: 'openid profile email' }),
    ).toThrow(/non-OIDC scope/i);
  });

  it('reads MSAL, notebook, and release profile config', () => {
    expect(msalConfigFromEnv(full)).toEqual({
      clientId: '22222222-2222-4222-8222-222222222222',
      authority: 'https://login.microsoftonline.com/11111111-1111-4111-8111-111111111111',
    });
    expect(notebookIdFromEnv(full)).toBe('nb-1');
    expect(notebookIdFromEnv({ ...full, VITE_NOTEBOOK_ID: undefined })).toBeUndefined();
    expect(releaseProfileFromEnv(full)).toBe('internal-alpha-word-excel');
    expect(widgetFromEnv({ ...full, VITE_GE_WIDGET_SERVER_TOKEN: undefined })).toEqual({
      configId: '33333333-3333-4333-8333-333333333333',
    });
  });

  it('accepts full skill resource names without expanding them', () => {
    const skill =
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/m365-surface-commander';
    expect(skillsFromEnv({ ...full, VITE_GE_SKILL_IDS: skill })).toEqual([skill]);
  });

  it('parses label-bound skill resources into mounted skills and GE mention markers', () => {
    const skill =
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129';
    const env = { ...full, VITE_GE_SKILL_IDS: `m365-surface-commander=${skill}` };
    expect(skillsFromEnv(env)).toEqual([skill]);
    expect(skillMentionsFromEnv(env)).toEqual([
      { label: 'm365-surface-commander', uri: '7404511736383961129' },
    ]);
    expect(shellConfigFromEnv(env).skillMentions).toEqual([
      { label: 'm365-surface-commander', uri: '7404511736383961129' },
    ]);
  });

  it('parses route-bound skill resources for planner and command turns', () => {
    expect(plannerSkillsFromEnv(full)).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/17573173582293271726',
    ]);
    expect(plannerSkillMentionsFromEnv(full)).toEqual([
      { label: 'm365-command-planner', uri: '17573173582293271726' },
    ]);
    expect(commandSkillsFromEnv(full)).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129',
    ]);
    expect(commandSkillMentionsFromEnv(full)).toEqual([
      { label: 'm365-surface-commander', uri: '7404511736383961129' },
    ]);
  });

  it('infers GE mention labels for unlabelled route-bound planner and command skills', () => {
    const env = {
      ...full,
      VITE_GE_COMMAND_PLANNER_SKILL: '17573173582293271726',
      VITE_GE_SURFACE_COMMANDER_SKILL: '7404511736383961129',
    };
    expect(plannerSkillsFromEnv(env)).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/17573173582293271726',
    ]);
    expect(plannerSkillMentionsFromEnv(env)).toEqual([
      { label: 'm365-command-planner', uri: '17573173582293271726' },
    ]);
    expect(commandSkillsFromEnv(env)).toEqual([
      'projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129',
    ]);
    expect(commandSkillMentionsFromEnv(env)).toEqual([
      { label: 'm365-surface-commander', uri: '7404511736383961129' },
    ]);
  });

  it('throws a helpful error when a required var is missing', () => {
    const missing = { ...full, VITE_GCP_PROJECT: undefined };
    expect(() => assistantFromEnv(missing)).toThrow(/VITE_GCP_PROJECT/);
    expect(() => assistantFromEnv(missing)).toThrow(/\.env\.example/);
  });

  it('rejects confidential browser configuration', () => {
    expect(() =>
      shellConfigFromEnv({
        ...full,
        VITE_ENTRA_CLIENT_SECRET: 'nope',
      } as RawEnv & { VITE_ENTRA_CLIENT_SECRET: string }),
    ).toThrow(/must never be present/i);
  });

  it('rejects unknown VITE_* keys in production', () => {
    expect(() =>
      shellConfigFromEnv({
        ...full,
        VITE_GATEWAY_URL: 'https://old.example',
      } as RawEnv & { VITE_GATEWAY_URL: string }),
    ).toThrow(/Unknown production browser configuration key VITE_GATEWAY_URL/);
  });

  it('rejects malformed residency and identifiers', () => {
    expect(() => assistantFromEnv({ ...full, VITE_GCP_LOCATION: 'moon-1' })).toThrow();
    expect(() => msalConfigFromEnv({ ...full, VITE_ENTRA_CLIENT_ID: 'not-a-guid' })).toThrow();
    expect(() => shellConfigFromEnv({ ...full, VITE_GE_SKILL_IDS: '../bad' })).toThrow();
    expect(() =>
      shellConfigFromEnv({ ...full, VITE_GE_COMMAND_PLANNER_SKILL_SOURCE_SHA256: 'not-a-sha' }),
    ).toThrow();
  });

  it('rejects production proxy URLs with unsafe schemes, localhost, or credentials', () => {
    expect(() => shellConfigFromEnv({ ...full, VITE_PROXY_URL: 'http://proxy.example' })).toThrow(
      /https/i,
    );
    expect(() =>
      shellConfigFromEnv({ ...full, VITE_PROXY_URL: 'https://user:pass@proxy.example' }),
    ).toThrow(/credentials/i);
    expect(() => shellConfigFromEnv({ ...full, VITE_PROXY_URL: 'https://localhost:8080' })).toThrow(
      /localhost/i,
    );
  });

  it('allows localhost proxy only outside production', () => {
    const dev = {
      ...full,
      MODE: 'development',
      PROD: false,
      VITE_PROXY_URL: 'http://localhost:8080',
    };
    expect(shellConfigFromEnv(dev).proxyUrl).toBe('http://localhost:8080');
  });
});
