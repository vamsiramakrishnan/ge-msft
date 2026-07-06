import { describe, expect, it, vi } from 'vitest';
import {
  DiscoveryCatalogClient,
  applyCatalogSelection,
  defaultCatalogSelection,
  groupDataStoresByConnector,
  type GeminiCatalogConnector,
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
      if (url.includes('/dataStores?')) return json({ dataStores: [] });
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
      if (url.includes('/dataStores?')) return json({ dataStores: [] });
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
      if (url.includes('/dataStores?'))
        return new Response('should not be called', { status: 500 });
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

    expect(fetchImpl.mock.calls.some((call) => String(call[0]).includes('/dataStores?'))).toBe(
      false,
    );
    expect(catalog.dataStores).toEqual([
      {
        name: 'projects/proj/locations/global/collections/default_collection/dataStores/msft-onedrive-fed_1779469629030_file',
        id: 'msft-onedrive-fed_1779469629030_file',
        label: 'OneDrive files',
        type: 'GENERIC',
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

  it('pages through collections.list and keeps only collections with a data connector', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (!url.pathname.endsWith('/collections')) return new Response('not found', { status: 404 });
      if (!url.searchParams.get('pageToken')) {
        return json({
          collections: [
            {
              name: 'projects/proj/locations/global/collections/default_collection',
              displayName: 'Default',
            },
            {
              name: 'projects/proj/locations/global/collections/sharepoint-connector',
              displayName: 'SharePoint',
              dataConnector: {
                dataSource: 'sharepoint',
                state: 'ACTIVE',
                connectorModes: ['FEDERATED'],
                lastSyncTime: '2026-07-01T10:00:00Z',
                entities: [{ entityName: 'site' }, { entityName: 'file' }],
              },
            },
          ],
          nextPageToken: 'col-page-2',
        });
      }
      return json({
        collections: [
          {
            name: 'projects/proj/locations/global/collections/jira-connector',
            dataConnector: {
              dataSource: 'jira',
              state: 'FAILED',
              errors: [{ message: 'credential expired' }],
              blockingReasons: ['ALLOWLIST_REQUIRED'],
            },
          },
        ],
      });
    });

    const connectors = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listConnectors();

    expect(fetchImpl.mock.calls.map((call) => String(call[0]))).toEqual([
      'https://discoveryengine.googleapis.com/v1alpha/projects/proj/locations/global/collections?pageSize=100',
      'https://discoveryengine.googleapis.com/v1alpha/projects/proj/locations/global/collections?pageSize=100&pageToken=col-page-2',
    ]);
    expect(connectors).toEqual([
      {
        name: 'projects/proj/locations/global/collections/sharepoint-connector',
        id: 'sharepoint-connector',
        label: 'SharePoint',
        source: 'sharepoint',
        state: 'active',
        modes: ['FEDERATED'],
        lastSyncTime: '2026-07-01T10:00:00Z',
        entities: ['site', 'file'],
      },
      {
        name: 'projects/proj/locations/global/collections/jira-connector',
        id: 'jira-connector',
        label: 'jira-connector',
        source: 'jira',
        state: 'failed',
        errorCount: 1,
        blockingReasons: ['ALLOWLIST_REQUIRED'],
      },
    ]);
  });

  it('maps INITIALIZATION_FAILED to failed and unrecognized connector states to unknown', async () => {
    const fetchImpl = vi.fn(async () =>
      json({
        collections: [
          {
            name: 'projects/proj/locations/global/collections/init-failed',
            dataConnector: { dataSource: 'jira', state: 'INITIALIZATION_FAILED' },
          },
          {
            name: 'projects/proj/locations/global/collections/brand-new',
            dataConnector: { dataSource: 'box', state: 'SOME_FUTURE_STATE' },
          },
        ],
      }),
    );

    const connectors = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listConnectors();

    expect(connectors.map((c) => [c.id, c.state])).toEqual([
      ['init-failed', 'failed'],
      ['brand-new', 'unknown'],
    ]);
  });

  it('degrades to an empty connector list with a warning when collections.list is denied', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url.includes('/agents?')) return json({ agents: [] });
      if (url.includes('/dataStores?')) {
        return json({
          dataStores: [
            {
              name: 'projects/proj/locations/global/collections/default_collection/dataStores/ds-1',
              displayName: 'One',
            },
          ],
        });
      }
      if (url.includes('/collections?')) {
        return new Response(
          JSON.stringify({
            error: { code: 403, status: 'PERMISSION_DENIED', message: 'denied' },
          }),
          { status: 403 },
        );
      }
      return new Response('not found', { status: 404 });
    });

    const catalog = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listCatalog();

    expect(catalog.connectors).toEqual([]);
    expect(catalog.dataStores.map((store) => store.id)).toEqual(['ds-1']);
    expect(
      catalog.warnings?.some((warning) => warning.includes('discoveryengine.collections.list')),
    ).toBe(true);
  });

  it('follows nextPageToken across admin dataStores pages', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (!url.pathname.endsWith('/dataStores')) return new Response('not found', { status: 404 });
      if (!url.searchParams.get('pageToken')) {
        return json({
          dataStores: [{ name: `${collectionParent()}/dataStores/ds-1`, displayName: 'One' }],
          nextPageToken: 'ds-page-2',
        });
      }
      expect(url.searchParams.get('pageToken')).toBe('ds-page-2');
      return json({
        dataStores: [{ name: `${collectionParent()}/dataStores/ds-2`, displayName: 'Two' }],
      });
    });

    const stores = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listAdminDataStores();

    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(String(fetchImpl.mock.calls[0]?.[0])).toContain('/dataStores?pageSize=100');
    expect(String(fetchImpl.mock.calls[1]?.[0])).toContain('pageToken=ds-page-2');
    expect(stores.map((store) => store.id)).toEqual(['ds-1', 'ds-2']);
  });

  it('follows nextPageToken across admin agent pages', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (!url.pathname.endsWith('/agents')) return new Response('not found', { status: 404 });
      if (!url.searchParams.get('pageToken')) {
        return json({
          agents: [{ name: `${agentParent()}/agents/1`, displayName: 'skill-a' }],
          nextPageToken: 'agents-page-2',
        });
      }
      expect(url.searchParams.get('pageToken')).toBe('agents-page-2');
      return json({ agents: [{ name: `${agentParent()}/agents/2`, displayName: 'skill-b' }] });
    });

    const skills = await new DiscoveryCatalogClient(
      tokenSource,
      cfg,
      fetchImpl as never,
    ).listAdminSkills();

    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(String(fetchImpl.mock.calls[0]?.[0])).toContain('/agents?pageSize=100');
    expect(String(fetchImpl.mock.calls[1]?.[0])).toContain('pageToken=agents-page-2');
    expect(skills.map((skill) => skill.label)).toEqual(['skill-a', 'skill-b']);
  });

  it('follows widget agent-view pages via listAvailableAgentViewsRequest.pageToken', async () => {
    const fetchImpl = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      if (!url.includes('/widgetListAvailableAgentViews')) {
        return new Response('not found', { status: 404 });
      }
      const request = JSON.parse(String(init?.body)).listAvailableAgentViewsRequest;
      if (request.agentOrigin === 'GOOGLE') return json({ agentViews: [] });
      if (!request.pageToken) {
        return json({
          agentViews: [
            { agent: { name: `${agentParent()}/agents/1`, displayName: 'skill-a' }, uri: '1' },
          ],
          nextPageToken: 'views-page-2',
        });
      }
      expect(request.pageToken).toBe('views-page-2');
      return json({
        agentViews: [
          { agent: { name: `${agentParent()}/agents/2`, displayName: 'skill-b' }, uri: '2' },
        ],
      });
    });

    const skills = await new DiscoveryCatalogClient(
      tokenSource,
      { ...cfg, widget: { configId: 'test-widget-config-id' } },
      fetchImpl as never,
    ).listWidgetSkills();

    const userBodies = fetchImpl.mock.calls
      .map((call) => JSON.parse(String((call[1] as RequestInit).body)))
      .filter((body) => body.listAvailableAgentViewsRequest.agentOrigin === 'USER');
    expect(userBodies).toHaveLength(2);
    expect(userBodies[1].listAvailableAgentViewsRequest.pageToken).toBe('views-page-2');
    expect(skills.map((skill) => skill.label)).toEqual(['skill-a', 'skill-b']);
  });

  it('groups data stores under their connector collection and leaves the rest ungrouped', () => {
    const connector: GeminiCatalogConnector = {
      name: 'projects/proj/locations/global/collections/sharepoint-connector',
      id: 'sharepoint-connector',
      label: 'SharePoint',
      source: 'sharepoint',
      state: 'active',
    };
    const grouped = {
      name: 'projects/proj/locations/global/collections/sharepoint-connector/dataStores/sp-files',
      id: 'sp-files',
      label: 'SP files',
    };
    const loose = {
      name: 'projects/proj/locations/global/collections/default_collection/dataStores/ds-web',
      id: 'ds-web',
      label: 'Web',
    };

    expect(groupDataStoresByConnector([connector], [grouped, loose])).toEqual({
      groups: [{ connector, stores: [grouped] }],
      ungrouped: [loose],
    });
  });
});

function collectionParent(): string {
  return 'projects/proj/locations/global/collections/default_collection';
}

function agentParent(): string {
  return (
    'projects/proj/locations/global/collections/default_collection' +
    '/engines/test-engine/assistants/default_assistant'
  );
}
