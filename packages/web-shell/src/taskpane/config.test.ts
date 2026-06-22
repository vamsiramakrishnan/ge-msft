import { describe, it, expect } from 'vitest';
import {
  assistantFromEnv,
  wifFromEnv,
  shellConfigFromEnv,
  authOptionsFromEnv,
  msalConfigFromEnv,
  notebookIdFromEnv,
  type RawEnv,
} from './config.js';

const full: RawEnv = {
  VITE_GCP_PROJECT: 'proj-1',
  VITE_GCP_LOCATION: 'eu',
  VITE_GE_ENGINE: 'engine-1',
  VITE_GE_COLLECTION: 'col-1',
  VITE_GE_ASSISTANT: 'assist-1',
  VITE_GE_MODEL_ID: 'gemini-x',
  VITE_WIF_POOL_ID: 'pool-1',
  VITE_WIF_PROVIDER_ID: 'prov-1',
  VITE_WIF_SCOPE: 'https://www.googleapis.com/auth/cloud-platform',
  VITE_WIF_USER_PROJECT: 'billing-1',
  VITE_PROXY_URL: 'https://proxy.example.com',
  VITE_ENTRA_CLIENT_ID: 'client-1',
  VITE_ENTRA_AUTHORITY: 'https://login.microsoftonline.com/tenant',
  VITE_WIF_ID_TOKEN_SCOPES: 'openid profile',
  VITE_GRAPH_SCOPES: 'User.Read Files.Read.All',
  VITE_NOTEBOOK_ID: 'nb-1',
};

describe('config from env', () => {
  it('builds the assistant path with optional fields', () => {
    expect(assistantFromEnv(full)).toEqual({
      project: 'proj-1',
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
      userProject: 'billing-1',
    });
  });

  it('assembles the full shell config (no secrets)', () => {
    const cfg = shellConfigFromEnv(full);
    expect(cfg.assistant.project).toBe('proj-1');
    expect(cfg.modelId).toBe('gemini-x');
    expect(cfg.proxyUrl).toBe('https://proxy.example.com');
    expect(JSON.stringify(cfg)).not.toMatch(/secret/i);
  });

  it('parses space/comma scopes; falls back to defaults', () => {
    expect(authOptionsFromEnv(full).idTokenScopes).toEqual(['openid', 'profile']);
    expect(authOptionsFromEnv(full).graphScopes).toEqual(['User.Read', 'Files.Read.All']);
    const noScopes = { ...full, VITE_GRAPH_SCOPES: undefined };
    expect(authOptionsFromEnv(noScopes).graphScopes).toEqual(['User.Read']);
  });

  it('reads MSAL + notebook config', () => {
    expect(msalConfigFromEnv(full)).toEqual({
      clientId: 'client-1',
      authority: 'https://login.microsoftonline.com/tenant',
    });
    expect(notebookIdFromEnv(full)).toBe('nb-1');
    expect(notebookIdFromEnv({ ...full, VITE_NOTEBOOK_ID: '' })).toBeUndefined();
  });

  it('throws a helpful error when a required var is missing', () => {
    const missing = { ...full, VITE_GCP_PROJECT: undefined };
    expect(() => assistantFromEnv(missing)).toThrow(/VITE_GCP_PROJECT/);
    expect(() => assistantFromEnv(missing)).toThrow(/\.env\.example/);
  });
});
