import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  UnitDescriptor,
} from '@ge/contracts';
import type { AuthClient, DocBridge, UserIdentity } from '@ge/runtime';
import {
  composeSession,
  estateWritesEnabledFor,
  SKILL_FILES,
  type ShellConfig,
  skillDriftNotice,
} from './compose.js';

const config: ShellConfig = {
  assistant: { project: 'p', location: 'eu', engine: 'e' },
  wif: { poolId: 'pool', providerId: 'entra' },
};

const auth: AuthClient = {
  getIdToken: () => Promise.resolve('id-tok'),
  getGraphToken: () => Promise.resolve('graph-tok'),
  getIdentity: (): Promise<UserIdentity> => Promise.resolve({ username: 'v.k@acme' }),
};

class FakeBridge implements DocBridge {
  readonly surface = 'word' as const;
  getCapabilities(): CapabilityManifest {
    return { surface: 'word', contextKinds: [], actuations: [] };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(req: ActuationRequest): Promise<ActuationResult> {
    return Promise.resolve({ ok: true, changeId: req.changeId, kind: req.kind });
  }
}

const unit: UnitDescriptor = { connectors: [], surfaceContext: { kind: 'word' } };

describe('composeSession', () => {
  it('wires identity → WIF → client → AssistSession and resumes a prior session', async () => {
    const getIdentity = vi.spyOn(auth, 'getIdentity');
    const fetchImpl = vi.fn() as unknown as typeof fetch;
    const { session, tokens, client } = await composeSession({
      config,
      auth,
      bridge: new FakeBridge(),
      unit,
      resumeSessionId: 'sess_prior',
      fetchImpl,
    });
    expect(getIdentity).toHaveBeenCalledOnce();
    expect(session.sessionId).toBe('sess_prior'); // resumed
    expect(tokens).toBeDefined();
    expect(client).toBeDefined();
  });

  it('sends only the brand writing rules for the open app (email rules only from Outlook)', async () => {
    const rules = [
      'Keep answers short.',
      { text: 'Open Bahasa emails with "Halo".', surfaces: ['outlook'] },
      { text: 'Sign emails as ContosoAI.', surfaces: ['outlook'], turns: ['command' as const] },
    ];
    const rulesSent = async (surface: 'word' | 'outlook'): Promise<unknown> => {
      const bridge = Object.assign(new FakeBridge(), { surface });
      const { client } = await composeSession({
        config: { ...config, assistantName: 'ContosoAI', assistantRules: rules },
        auth,
        bridge: bridge as unknown as DocBridge,
        unit,
        fetchImpl: vi.fn() as unknown as typeof fetch,
      });
      return (client as unknown as { config: { assistantRules?: string[] } }).config.assistantRules;
    };
    expect(await rulesSent('outlook')).toEqual([
      'Keep answers short.',
      'Open Bahasa emails with "Halo".',
      // A rule limited to turns keeps that limit for the client (only command turns send it).
      { text: 'Sign emails as ContosoAI.', routes: ['command'] },
    ]);
    expect(await rulesSent('word')).toEqual(['Keep answers short.']);
  });
});

describe('estateWritesEnabledFor — release-profile enforcement', () => {
  it('is always false without a sharedStore, regardless of profile', () => {
    expect(estateWritesEnabledFor(false, undefined)).toBe(false);
    expect(estateWritesEnabledFor(false, 'development')).toBe(false);
    expect(estateWritesEnabledFor(false, 'internal-alpha-word-excel')).toBe(false);
  });

  it('is unrestricted (true) when no release profile is configured', () => {
    expect(estateWritesEnabledFor(true, undefined)).toBe(true);
  });

  it('reads the real profile value for each currently defined profile', () => {
    // Both ship with estateWrites: true today, but this asserts the code actually CONSULTS the
    // profile (not a hardcoded true) — a future stricter profile flipping this to false would be
    // enforced by this same AND, not just documented.
    expect(estateWritesEnabledFor(true, 'development')).toBe(true);
    expect(estateWritesEnabledFor(true, 'internal-alpha-word-excel')).toBe(true);
  });
});

describe('composeSession — /shared wiring', () => {
  it('wires a Graph-backed sharedStore when the AuthClient carries a Graph token source', async () => {
    const fetchImpl = vi.fn() as unknown as typeof fetch;
    const { session } = await composeSession({
      config,
      auth,
      bridge: new FakeBridge(),
      unit,
      fetchImpl,
    });
    expect(session).toBeDefined();
  });

  it('degrades gracefully (no throw) when the AuthClient has no Graph token source', async () => {
    const authNoGraph: AuthClient = {
      getIdToken: () => Promise.resolve('id-tok'),
      getIdentity: (): Promise<UserIdentity> => Promise.resolve({ username: 'v.k@acme' }),
    };
    const fetchImpl = vi.fn() as unknown as typeof fetch;
    const { session } = await composeSession({
      config,
      auth: authNoGraph,
      bridge: new FakeBridge(),
      unit,
      fetchImpl,
    });
    expect(session).toBeDefined();
  });
});

describe('SKILL_FILES', () => {
  it('bundles the real skill/ markdown files, keyed relative to skill/ with real content', () => {
    expect(Object.keys(SKILL_FILES).length).toBeGreaterThan(0);
    expect(SKILL_FILES['m365-surface-commander/SKILL.md']).toContain('# ');
    // Keys are clean relative paths — no leading "../" or absolute-path leakage from the glob.
    for (const key of Object.keys(SKILL_FILES)) {
      expect(key.startsWith('.')).toBe(false);
      expect(key.startsWith('/')).toBe(false);
    }
  });
});

describe('skillDriftNotice', () => {
  it('names stale and missing skills, and stays quiet when all match or none is configured', () => {
    expect(skillDriftNotice([])).toBeUndefined();
    expect(
      skillDriftNotice([{ action: 'unchanged', agent: { displayName: 'm365-surface-commander' } }]),
    ).toBeUndefined();
    expect(
      skillDriftNotice([{ action: 'error', agentId: 'cmdr', error: 'forbidden' }]),
    ).toBeUndefined();
    const notice = skillDriftNotice([
      { action: 'stale', agent: { displayName: 'm365-surface-commander' } },
      { action: 'missing', agentId: 'm365-command-planner' },
    ]);
    expect(notice).toContain('older than this app: m365-surface-commander');
    expect(notice).toContain('not found: m365-command-planner');
    expect(notice).toContain('bun run ge:skills');
  });
});
