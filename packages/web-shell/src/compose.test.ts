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
import { composeSession, type ShellConfig } from './compose.js';

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
});
