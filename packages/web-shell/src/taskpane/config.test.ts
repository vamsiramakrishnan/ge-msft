import { describe, it, expect, vi, beforeEach, afterEach, type MockInstance } from 'vitest';
import {
  assistantFromEnv,
  wifFromEnv,
  shellConfigFromEnv,
  authOptionsFromEnv,
  msalConfigFromEnv,
  notebookIdFromEnv,
  releaseProfileFromEnv,
  resolveRuntimeEnv,
  commandSkillMentionsFromEnv,
  commandSkillsFromEnv,
  plannerSkillMentionsFromEnv,
  plannerSkillsFromEnv,
  skillMentionsFromEnv,
  skillsFromEnv,
  widgetFromEnv,
  type RawEnv,
} from './config.js';

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
  VITE_GE_SURFACE_COMMANDER_SKILL:
    'm365-surface-commander=projects/proj-12345/locations/eu/collections/col-1/engines/engine-1/assistants/assist-1/agents/7404511736383961129',
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

// ---- Runtime tenant-config bootstrap (ADR-0009 first slice) ----

const ORIGIN = 'https://addin.contoso.example';

/** Stub the browser globals resolveRuntimeEnv reads: window.location and fetch. */
function stubBrowser(search: string): ReturnType<typeof vi.fn> {
  vi.stubGlobal('window', { location: { origin: ORIGIN, search } });
  const fetchMock = vi.fn();
  vi.stubGlobal('fetch', fetchMock);
  return fetchMock;
}

/** Minimal Response stand-in: only the members resolveRuntimeEnv touches. */
function jsonResponse(body: unknown, ok = true): Response {
  return { ok, status: ok ? 200 : 404, json: () => Promise.resolve(body) } as unknown as Response;
}

describe('resolveRuntimeEnv (runtime tenant config, ADR-0009 first slice)', () => {
  let warn: MockInstance<Parameters<typeof console.warn>, void>;

  beforeEach(() => {
    warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });

  afterEach(() => {
    warn.mockRestore();
    vi.unstubAllGlobals();
  });

  it('fetches /config/<cfg>.json for a valid ?cfg= and overlays its values', async () => {
    const fetchMock = stubBrowser('?host=word&cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 'tenant-engine' }));
    const resolved = await resolveRuntimeEnv(full);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock).toHaveBeenCalledWith(
      `${ORIGIN}/config/contoso-prod.json`,
      expect.objectContaining({ cache: 'no-store', signal: expect.any(AbortSignal) }),
    );
    expect(resolved.VITE_GE_ENGINE).toBe('tenant-engine');
    // Untouched keys survive the merge, and the merged env still parses through parseEnv.
    expect(resolved.VITE_GCP_PROJECT).toBe('proj-12345');
    expect(assistantFromEnv(resolved).engine).toBe('tenant-engine');
  });

  it('falls back to /config/default.json when the cfg-specific file is missing', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock
      .mockResolvedValueOnce(jsonResponse(null, false)) // 404 on contoso-prod.json
      .mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 'default-engine' }));
    const resolved = await resolveRuntimeEnv(full);
    expect(fetchMock).toHaveBeenNthCalledWith(
      2,
      `${ORIGIN}/config/default.json`,
      expect.anything(),
    );
    expect(resolved.VITE_GE_ENGINE).toBe('default-engine');
  });

  it('ignores an invalid ?cfg= value instead of interpolating it', async () => {
    const fetchMock = stubBrowser(`?cfg=${encodeURIComponent('../../../etc/passwd')}`);
    fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 'default-engine' }));
    const resolved = await resolveRuntimeEnv(full);
    // The traversal-shaped value never reaches a URL; only the default candidate is fetched.
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]?.[0]).toBe(`${ORIGIN}/config/default.json`);
    expect(resolved.VITE_GE_ENGINE).toBe('default-engine');
  });

  it('cannot fetch cross-origin by construction: every candidate URL keeps our origin', async () => {
    const fetchMock = stubBrowser(`?cfg=${encodeURIComponent('evil.example/steal')}`);
    fetchMock.mockResolvedValue(jsonResponse(null, false));
    await resolveRuntimeEnv(full);
    expect(fetchMock.mock.calls.length).toBeGreaterThan(0);
    for (const call of fetchMock.mock.calls) {
      expect(new URL(call[0] as string).origin).toBe(ORIGIN);
      expect(call[0]).toMatch(new RegExp(`^${ORIGIN}/config/[a-z0-9-]+\\.json$`));
    }
  });

  it('rejects a document with unknown keys wholesale (fail closed, no partial salvage)', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ VITE_GE_ENGINE: 'tenant-engine', VITE_GATEWAY_URL: 'https://old.example' }),
    );
    const resolved = await resolveRuntimeEnv(full);
    expect(resolved).toEqual(full); // even the known key is NOT applied
    expect(fetchMock).toHaveBeenCalledTimes(1); // an invalid doc does not fall through to default
    expect(warn).toHaveBeenCalledWith(expect.stringMatching(/unknown key/));
    // Untrusted key text must not be echoed to the console (only a count is logged).
    expect(warn.mock.calls.map((c) => String(c[0])).join(' ')).not.toContain('VITE_GATEWAY_URL');
  });

  it('ignores an encoded-traversal or uppercase ?cfg= value after URL decoding', async () => {
    for (const cfg of ['%2e%2e%2fconfig', 'ACME', 'a'.repeat(65)]) {
      const fetchMock = stubBrowser(`?cfg=${cfg}`);
      fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 'default-engine' }));
      await resolveRuntimeEnv(full);
      // URLSearchParams decodes BEFORE CFG_NAME runs, so '..' and '/' can never survive into a URL.
      expect(fetchMock).toHaveBeenCalledTimes(1);
      expect(fetchMock.mock.calls[0]?.[0]).toBe(`${ORIGIN}/config/default.json`);
      vi.unstubAllGlobals();
    }
  });

  it('rejects a document containing secret-shaped values wholesale', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ VITE_PROXY_URL: 'https://proxy.example/?access_token=eyJhbGci' }),
    );
    const resolved = await resolveRuntimeEnv(full);
    expect(resolved).toEqual(full);
    expect(warn).toHaveBeenCalledWith(expect.stringMatching(/secret-shaped/));
    // The warning must not echo the value itself.
    expect(warn.mock.calls.map((c) => String(c[0])).join(' ')).not.toContain('eyJhbGci');
  });

  it('rejects non-object and array documents', async () => {
    for (const body of ['just-a-string', ['VITE_GE_ENGINE'], 42, null]) {
      const fetchMock = stubBrowser('?cfg=contoso-prod');
      fetchMock.mockResolvedValueOnce(jsonResponse(body));
      expect(await resolveRuntimeEnv(full)).toEqual(full);
      vi.unstubAllGlobals();
    }
    expect(warn).toHaveBeenCalledTimes(4);
  });

  it('rejects documents with non-string values', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 42 }));
    expect(await resolveRuntimeEnv(full)).toEqual(full);
    expect(warn).toHaveBeenCalledWith(expect.stringMatching(/non-string/));
  });

  it('rejects documents with empty-string values (deletion, not selection)', async () => {
    // An empty string would UNSET a build-time value through parseEnv's falsy-spread (or
    // fail-stop boot on a required key), so it must reject the whole document.
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ VITE_GE_ENGINE: 'tenant-engine', VITE_GE_RELEASE_PROFILE: '  ' }),
    );
    const resolved = await resolveRuntimeEnv(full);
    expect(resolved).toEqual(full); // the release profile survives; the engine is NOT applied
    expect(warn).toHaveBeenCalledWith(expect.stringMatching(/empty values/));
  });

  it('resolves within the timeout when the server never answers (hung headers)', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockImplementation(
      (_url: string, init?: RequestInit) =>
        new Promise((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => reject(new Error('aborted')));
        }),
    );
    vi.useFakeTimers();
    try {
      const pending = resolveRuntimeEnv(full);
      await vi.advanceTimersByTimeAsync(4001); // aborts the cfg-specific candidate
      await vi.advanceTimersByTimeAsync(4001); // aborts the default.json fallback
      expect(await pending).toEqual(full);
      expect(fetchMock).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it('resolves within the timeout when the body never completes (hung json())', async () => {
    // Regression: the abort timer must cover the BODY read, not just the headers — a server that
    // answers headers and then trickles the body forever must not wedge boot.
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockImplementation((_url: string, init?: RequestInit) =>
      Promise.resolve({
        ok: true,
        status: 200,
        json: () =>
          new Promise((_resolve, reject) => {
            init?.signal?.addEventListener('abort', () => reject(new Error('aborted')));
          }),
      } as unknown as Response),
    );
    vi.useFakeTimers();
    try {
      const pending = resolveRuntimeEnv(full);
      await vi.advanceTimersByTimeAsync(4001);
      await vi.advanceTimersByTimeAsync(4001);
      expect(await pending).toEqual(full);
      expect(fetchMock).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });

  it('rejects malformed JSON without falling through to a weaker candidate', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce({
      ok: true,
      status: 200,
      json: () => Promise.reject(new SyntaxError('bad json')),
    } as unknown as Response);
    expect(await resolveRuntimeEnv(full)).toEqual(full);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(warn).toHaveBeenCalledWith(expect.stringMatching(/not valid JSON/));
  });

  it('returns the base env unchanged when every candidate fails on the network', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockRejectedValue(new TypeError('network down'));
    expect(await resolveRuntimeEnv(full)).toEqual(full);
    expect(fetchMock).toHaveBeenCalledTimes(2); // cfg-specific candidate, then default.json
  });

  it('lets fetched values override base values while keeping the rest', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(
      jsonResponse({ VITE_GE_ENGINE: 'tenant-engine', VITE_GCP_LOCATION: 'us' }),
    );
    const resolved = await resolveRuntimeEnv(full);
    expect(resolved.VITE_GE_ENGINE).toBe('tenant-engine');
    expect(resolved.VITE_GCP_LOCATION).toBe('us');
    expect(resolved.VITE_WIF_POOL_ID).toBe('pool-1');
    expect(resolved.MODE).toBe('production'); // build-time flags are not runtime-overridable
  });

  it('still fails parseEnv when a required key is missing from both base and document', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_GE_ENGINE: 'tenant-engine' }));
    const { VITE_GCP_PROJECT: _p, ...missingBase } = full;
    const resolved = await resolveRuntimeEnv(missingBase);
    expect(() => assistantFromEnv(resolved)).toThrow(/VITE_GCP_PROJECT/);
  });

  it('does not weaken production checks on the merged result', async () => {
    const fetchMock = stubBrowser('?cfg=contoso-prod');
    fetchMock.mockResolvedValueOnce(jsonResponse({ VITE_PROXY_URL: 'http://insecure.example' }));
    const resolved = await resolveRuntimeEnv(full);
    expect(resolved.VITE_PROXY_URL).toBe('http://insecure.example');
    expect(() => shellConfigFromEnv(resolved)).toThrow(/https/i);
  });

  it('returns the base env when no window is available', async () => {
    vi.stubGlobal('fetch', vi.fn());
    expect(await resolveRuntimeEnv(full)).toEqual(full);
  });
});
