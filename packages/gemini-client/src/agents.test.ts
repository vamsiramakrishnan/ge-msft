import { describe, expect, it } from 'vitest';
import type { GeminiClientConfig } from './config.js';
import type { TokenSource } from './stream-assist.js';
import { ensureSkillAgent, getAgent, listAvailableAgentViews, listSkillAgents } from './agents.js';

const cfg: GeminiClientConfig = {
  assistant: { project: 'p', location: 'global', engine: 'e', assistant: 'default_assistant' },
};
const tokens: TokenSource = { getAccessToken: async () => 'tok' };

const AGENT_BASE =
  'https://discoveryengine.googleapis.com/v1alpha/projects/p/locations/global' +
  '/collections/default_collection/engines/e/assistants/default_assistant/agents';

/** Records requests and replies from a scripted route table keyed by `${method} ${pathIncludes}`. */
function mockFetch(
  routes: Array<{ match: (url: string, init?: RequestInit) => boolean; res: () => Response }>,
) {
  const calls: Array<{ url: string; method: string; body?: string }> = [];
  const fetchImpl = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = String(input);
    calls.push({ url, method: init?.method ?? 'GET', body: init?.body as string | undefined });
    const route = routes.find((r) => r.match(url, init));
    if (!route) return new Response('{}', { status: 404 });
    return route.res();
  }) as unknown as typeof fetch;
  return { fetchImpl, calls };
}

const json = (obj: unknown, status = 200) =>
  new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json' } });

describe('getAgent', () => {
  it('returns null on 404', async () => {
    const { fetchImpl } = mockFetch([
      { match: () => true, res: () => new Response('not found', { status: 404 }) },
    ]);
    expect(await getAgent(cfg, 'x', { tokens, fetchImpl })).toBeNull();
  });

  it('parses an existing agent', async () => {
    const { fetchImpl } = mockFetch([
      {
        match: (u) => u.endsWith('/agents/x'),
        res: () => json({ name: `${AGENT_BASE}/x`, displayName: 'X' }),
      },
    ]);
    const a = await getAgent(cfg, 'x', { tokens, fetchImpl });
    expect(a?.displayName).toBe('X');
  });
});

describe('ensureSkillAgent (warm-up)', () => {
  const input = {
    agentId: 'skill1',
    displayName: 'Skill One',
    description: 'does things',
    instruction: 'be helpful',
    revision: 'sha-abc',
  };

  it('creates when missing and stamps the revision into the description', async () => {
    const { fetchImpl, calls } = mockFetch([
      {
        match: (u, i) => (i?.method ?? 'GET') === 'GET',
        res: () => new Response('nope', { status: 404 }),
      },
      {
        match: (u, i) => i?.method === 'POST',
        res: () => json({ name: `${AGENT_BASE}/skill1`, displayName: 'Skill One' }),
      },
    ]);
    const r = await ensureSkillAgent(cfg, input, { tokens, fetchImpl });
    expect(r.action).toBe('created');
    const post = calls.find((c) => c.method === 'POST');
    expect(post?.url).toContain('agents?agentId=skill1');
    expect(JSON.parse(post!.body!).description).toBe('does things [rev:sha-abc]');
  });

  it('is a no-op when the stored revision matches (GET only, no write)', async () => {
    const { fetchImpl, calls } = mockFetch([
      {
        match: (u, i) => (i?.method ?? 'GET') === 'GET',
        res: () => json({ name: `${AGENT_BASE}/skill1`, description: 'does things [rev:sha-abc]' }),
      },
    ]);
    const r = await ensureSkillAgent(cfg, input, { tokens, fetchImpl });
    expect(r.action).toBe('unchanged');
    expect(calls.every((c) => c.method === 'GET')).toBe(true);
  });

  it('updates when the revision drifted', async () => {
    const { fetchImpl, calls } = mockFetch([
      {
        match: (u, i) => (i?.method ?? 'GET') === 'GET',
        res: () => json({ name: `${AGENT_BASE}/skill1`, description: 'does things [rev:OLD]' }),
      },
      { match: (u, i) => i?.method === 'PATCH', res: () => json({ name: `${AGENT_BASE}/skill1` }) },
    ]);
    const r = await ensureSkillAgent(cfg, input, { tokens, fetchImpl });
    expect(r.action).toBe('updated');
    const patch = calls.find((c) => c.method === 'PATCH');
    expect(patch?.url).toContain('updateMask=');
    expect(JSON.parse(patch!.body!).description).toBe('does things [rev:sha-abc]');
  });
});

describe('listAvailableAgentViews', () => {
  it('extracts ids from resource names and can filter to skill agents', async () => {
    const { fetchImpl } = mockFetch([
      {
        match: (_u, i) => i?.method === 'POST',
        res: () =>
          json({
            agentViews: [
              {
                name: `${AGENT_BASE}/3708`,
                displayName: 'm365-surface-commander',
                agentType: 'SKILL_AGENT',
                state: 'PRIVATE',
              },
              {
                name: `${AGENT_BASE}/deep_research`,
                displayName: 'Deep Research',
                agentType: 'MANAGED',
                state: 'ENABLED',
              },
            ],
          }),
      },
    ]);
    const all = await listAvailableAgentViews(cfg, { tokens, fetchImpl });
    expect(all.map((v) => v.id)).toEqual(['3708', 'deep_research']);
    const skills = await listSkillAgents(cfg, { tokens, fetchImpl });
    expect(skills.map((v) => v.displayName)).toEqual(['m365-surface-commander']);
  });
});
