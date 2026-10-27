import { describe, expect, it, vi } from 'vitest';
import { loadTeamsSdk, teamsBridgeOptions, type TeamsSdk } from './teams-sdk.js';

function sdk(overrides: Partial<TeamsSdk['app']> = {}): TeamsSdk {
  return {
    app: {
      initialize: () => Promise.resolve(),
      getContext: () => Promise.resolve({ chat: { id: '19:c' } }),
      ...overrides,
    },
    meeting: { registerMeetingEndHandler: vi.fn() },
  };
}

describe('loadTeamsSdk', () => {
  it('returns the initialized SDK', async () => {
    const s = sdk();
    expect(await loadTeamsSdk(() => Promise.resolve(s))).toBe(s);
  });

  it('gives up when initialization never resolves (not inside a Teams frame)', async () => {
    const stalled = sdk({ initialize: () => new Promise<void>(() => {}) });
    expect(await loadTeamsSdk(() => Promise.resolve(stalled), 10)).toBeUndefined();
  });

  it('gives up when the SDK cannot be loaded', async () => {
    expect(await loadTeamsSdk(() => Promise.reject(new Error('no module')))).toBeUndefined();
  });
});

describe('teamsBridgeOptions', () => {
  it('passes only the feature-detected slice the bridge uses', () => {
    const s = sdk({ openLink: vi.fn() });
    const options = teamsBridgeOptions(s);
    expect(options?.teams?.app?.openLink).toBeDefined();
    expect(options?.teams?.meeting?.registerMeetingEndHandler).toBeDefined();
    expect(options?.teams).not.toHaveProperty('chat');
    expect(teamsBridgeOptions(undefined)).toBeUndefined();
  });
});
