import { describe, expect, it, vi } from 'vitest';
import {
  DiscoveryCatalogClient,
  applyCatalogSelection,
  defaultCatalogSelection,
} from './catalog.js';
import type { GeminiClientConfig } from './config.js';

const cfg: GeminiClientConfig = {
  assistant: {
    project: 'proj',
    location: 'global',
    collection: 'default_collection',
    engine: 'test-engine',
    assistant: 'default_assistant',
  },
};

const tokenSource = {
  getAccessToken: () => Promise.resolve('goog-token'),
};

function json(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('DiscoveryCatalogClient', () => {
  it('lists user-visible skills through the widget available-agent-views endpoint', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes('/widgetListAvailableAgentViews')) {
        const body = JSON.parse(String(init?.body));
        const origin = body.listAvailableAgentViewsRequest.agentOrigin as string;
        return json({
          agentViews: [
            {
              agent: {
                name: `${agentParent()}/agents/${origin === 'USER' ? '17573173582293271726' : '7404511736383961129'}`,
                displayName: origin === 'USER' ? 'm365-command-planner' : 'm365-surface-commander',
              },
              uri: origin === 'USER' ? '17573173582293271726' : '7404511736383961129',
            },
          ],
        });
      }
      if (url.endsWith('/dataStores')) return json({ dataStores: [] });
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      {
        ...cfg,
        widget: {
          configId: 'test-widget-config-id',
          serverToken: 'test-widget-server-token',
        },
      },
      fetchImpl as never,
    ).listCatalog();

    const widgetCalls = fetchImpl.mock.calls.filter((call) =>
      String(call[0]).includes('/widgetListAvailableAgentViews'),
    );
    expect(widgetCalls).toHaveLength(2);
    const bodies = widgetCalls.map((call) => JSON.parse(String((call[1] as RequestInit).body)));
    expect(bodies.map((body) => body.listAvailableAgentViewsRequest.agentOrigin).sort()).toEqual([
      'GOOGLE',
      'USER',
    ]);
    expect(bodies[0]).toMatchObject({
      configId: 'test-widget-config-id',
      additionalParams: { token: '-', origin: 'ORIGIN_UNSPECIFIED' },
      listAvailableAgentViewsRequest: {
        pageSize: 200,
        filter: 'agent_type = SKILL_AGENT',
      },
    });
    expect((widgetCalls[0]?.[1] as RequestInit).headers).toMatchObject({
      Authorization: 'Bearer goog-token',
      'Content-Type': 'application/json',
      'x-server-token': 'test-widget-server-token',
    });
    expect(catalog.skills.map((s) => [s.label, s.suggestedRoute])).toEqual([
      ['m365-command-planner', 'planner'],
      ['m365-surface-commander', 'command'],
    ]);
  });

  it('re-exchanges the signed-in user token once when widget catalog auth expires', async () => {
    const invalidate = vi.fn();
    let userOriginCalls = 0;
    const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes('/widgetListAvailableAgentViews')) {
        const body = JSON.parse(String(init?.body));
        const origin = body.listAvailableAgentViewsRequest.agentOrigin as string;
        if (origin === 'USER') {
          userOriginCalls += 1;
          if (userOriginCalls === 1) return new Response('expired', { status: 401 });
          return json({
            agentViews: [
              {
                agent: {
                  name: `${agentParent()}/agents/17573173582293271726`,
                  displayName: 'm365-command-planner',
                },
                uri: '17573173582293271726',
              },
            ],
          });
        }
        return json({ agentViews: [] });
      }
      if (url.endsWith('/dataStores')) return json({ dataStores: [] });
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      { getAccessToken: () => Promise.resolve('goog-token'), invalidate },
      {
        ...cfg,
        widget: {
          configId: 'test-widget-config-id',
        },
      },
      fetchImpl as never,
    ).listCatalog();

    expect(invalidate).toHaveBeenCalledTimes(1);
    expect(userOriginCalls).toBe(2);
    expect(catalog.skills[0]?.label).toBe('m365-command-planner');
  });

  it('uses widget config lookup for connector discovery before admin dataStores.list', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes('/widgetListAvailableAgentViews')) return json({ agentViews: [] });
      if (url.includes('/widgetQueryAvailableConnectorNodes')) return json({ connectorNodes: [] });
      if (url.includes('/lookupWidgetConfig')) {
        expect(JSON.parse(String(init?.body))).toEqual({
          widgetConfigId: 'test-widget-config-id',
        });
        return json({
          widgetConfig: {
            toolsSpec: {
              vertexAiSearchSpec: {
                dataStoreSpecs: [
                  {
                    dataStore:
                      'collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
                    displayName: 'OneDrive files',
                    type: 'GENERIC',
                    dataConnector:
                      'projects/proj/locations/global/collections/msft-onedrive-fed_1779469629030/dataConnector',
                  },
                  {
                    dataStoreId: 'msft-outlook-fed_1779468500280_mail',
                    displayName: 'Outlook mail',
                  },
                ],
              },
            },
          },
        });
      }
      if (url.endsWith('/dataStores')) return new Response('should not be called', { status: 500 });
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      {
        ...cfg,
        widget: {
          configId: 'test-widget-config-id',
        },
      },
      fetchImpl as never,
    ).listCatalog();

    expect(fetchImpl.mock.calls.some((call) => String(call[0]).endsWith('/dataStores'))).toBe(
      false,
    );
    expect(catalog.dataStores).toEqual([
      {
        name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
        id: 'msft-onedrive-fed_1779469629030_file',
        label: 'OneDrive files',
        type: 'GENERIC',
        connectorId: 'msft-onedrive-fed_1779469629030',
        suggested: true,
      },
      {
        name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-outlook-fed_1779468500280_mail',
        id: 'msft-outlook-fed_1779468500280_mail',
        label: 'Outlook mail',
        suggested: true,
      },
    ]);
  });

  it('uses widget connector-node discovery before admin collections.list', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (url.includes('/lookupWidgetConfig')) {
        return json({
          widgetConfig: {
            toolsSpec: {
              vertexAiSearchSpec: {
                dataStoreSpecs: [
                  {
                    dataStore:
                      'collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
                    dataConnector:
                      'projects/proj/locations/global/collections/msft-onedrive-fed_1779469629030/dataConnector',
                  },
                ],
              },
            },
          },
        });
      }
      if (url.includes('/widgetQueryAvailableConnectorNodes')) {
        expect(JSON.parse(String(init?.body))).toMatchObject({
          configId: 'test-widget-config-id',
          queryAvailableConnectorNodesRequest: {
            parent: 'projects/proj/locations/global',
            dataConnectors: [
              'projects/proj/locations/global/collections/msft-onedrive-fed_1779469629030/dataConnector',
            ],
            engine:
              'projects/proj/locations/global/collections/default_collection/engines/test-engine',
            view: 'FULL',
            toolType: 'CONNECTOR_NODE',
          },
          location: 'global',
        });
        return json({
          connectorNodes: [
            {
              dataConnector:
                'projects/proj/locations/global/collections/msft-onedrive-fed_1779469629030/dataConnector',
              displayName: 'OneDrive',
              dataSource: 'onedrive',
              state: 'ACTIVE',
              lastSyncTime: '2026-07-01T10:00:00Z',
            },
          ],
        });
      }
      if (url.includes('/collections?'))
        return new Response('admin collections should not be called', { status: 500 });
      return new Response('not found', { status: 404 });
    });

    const connectors = await new DiscoveryCatalogClient(
      tokenSource,
      {
        ...cfg,
        widget: { configId: 'test-widget-config-id' },
      },
      fetchImpl as never,
    ).listConnectors();

    expect(fetchImpl.mock.calls.some((call) => String(call[0]).includes('/collections?'))).toBe(
      false,
    );
    expect(connectors).toEqual([
      {
        name: 'projects/proj/locations/global/collections/msft-onedrive-fed_1779469629030',
        id: 'msft-onedrive-fed_1779469629030',
        label: 'OneDrive',
        source: 'onedrive',
        state: 'active',
        lastSyncTime: '2026-07-01T10:00:00Z',
      },
    ]);
  });

  it('falls back to admin dataStores.list when widget config lookup has no connector metadata', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url.includes('/widgetListAvailableAgentViews')) return json({ agentViews: [] });
      if (url.includes('/lookupWidgetConfig')) return json({ widgetConfig: { displayName: 'GE' } });
      if (url.includes('/dataStores?')) {
        return json({
          dataStores: [
            {
              name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
              displayName: 'OneDrive files',
            },
          ],
        });
      }
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      {
        ...cfg,
        widget: { configId: 'test-widget-config-id' },
      },
      fetchImpl as never,
    ).listCatalog();

    expect(catalog.dataStores.map((store) => store.label)).toEqual(['OneDrive files']);
  });

  it('lists assistant skills and collection data stores as the signed-in user', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url.includes('/agents?')) {
        return json({
          agents: [
            {
              name: `${agentParent()}/agents/17573173582293271726`,
              displayName: 'm365-command-planner',
              state: 'PRIVATE',
            },
            {
              name: `${agentParent()}/agents/7404511736383961129`,
              displayName: 'm365-surface-commander',
            },
          ],
        });
      }
      if (url.includes('/dataStores?')) {
        return json({
          dataStores: [
            {
              name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
              displayName: 'OneDrive files',
              type: 'GENERIC',
            },
          ],
        });
      }
      if (url.includes('/collections?')) return json({ collections: [] });
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listCatalog();

    expect(fetchImpl).toHaveBeenCalledTimes(3);
    const firstCall = fetchImpl.mock.calls[0] as unknown as [RequestInfo | URL, RequestInit];
    const firstInit = firstCall[1];
    expect(firstInit.headers).toEqual({ Authorization: 'Bearer goog-token' });
    expect(catalog.skills.map((s) => [s.label, s.suggestedRoute])).toEqual([
      ['m365-command-planner', 'planner'],
      ['m365-surface-commander', 'command'],
    ]);
    expect(catalog.dataStores[0]).toMatchObject({
      label: 'OneDrive files',
      id: 'msft-onedrive-fed_1779469629030_file',
    });
  });

  it('turns catalog defaults into route-specific StreamAssist config', () => {
    const planner = {
      name: `${agentParent()}/agents/17573173582293271726`,
      id: '17573173582293271726',
      label: 'm365-command-planner',
      mention: { label: 'm365-command-planner', uri: '17573173582293271726' },
      suggestedRoute: 'planner' as const,
    };
    const commander = {
      name: `${agentParent()}/agents/7404511736383961129`,
      id: '7404511736383961129',
      label: 'm365-surface-commander',
      mention: { label: 'm365-surface-commander', uri: '7404511736383961129' },
      suggestedRoute: 'command' as const,
    };
    const dataStore = {
      name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-outlook-fed_1779468500280_mail',
      id: 'msft-outlook-fed_1779468500280_mail',
      label: 'Outlook mail',
      suggested: true,
    };
    const defaults = defaultCatalogSelection({
      skills: [planner, commander],
      dataStores: [dataStore],
      connectors: [],
    });

    expect(applyCatalogSelection(defaults)).toEqual({
      skills: [],
      skillMentions: [],
      plannerSkills: [planner.name],
      plannerSkillMentions: [planner.mention],
      commandSkills: [commander.name],
      commandSkillMentions: [commander.mention],
      dataStores: [dataStore.name],
    });
  });

  it('falls back to configured route skills when the user cannot list agents', async () => {
    const planner = `${agentParent()}/agents/17573173582293271726`;
    const commander = `${agentParent()}/agents/7404511736383961129`;
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url.includes('/agents?')) {
        return new Response(
          JSON.stringify({
            error: {
              code: 403,
              status: 'PERMISSION_DENIED',
              message: 'User does not have permission to list all of the agents.',
            },
          }),
          { status: 403 },
        );
      }
      if (url.includes('/dataStores?')) return json({ dataStores: [] });
      if (url.includes('/collections?')) return json({ collections: [] });
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      {
        ...cfg,
        plannerSkills: [planner],
        plannerSkillMentions: [{ label: 'm365-command-planner', uri: '17573173582293271726' }],
        commandSkills: [commander],
        commandSkillMentions: [{ label: 'm365-surface-commander', uri: '7404511736383961129' }],
      },
      fetchImpl as never,
    ).listCatalog();

    expect(catalog.skills.map((s) => [s.label, s.suggestedRoute])).toEqual([
      ['m365-command-planner', 'planner'],
      ['m365-surface-commander', 'command'],
    ]);
    expect(catalog.warnings?.[0]).toContain('discoveryengine.agents.list');
  });
});

function agentParent(): string {
  return (
    'projects/proj/locations/global/collections/default_collection' +
    '/engines/test-engine/assistants/default_assistant'
  );
}
