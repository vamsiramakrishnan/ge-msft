import { describe, expect, it } from 'vitest';
import type { GeminiClientConfig } from './config.js';
import type { TokenSource } from './stream-assist.js';
import {
  connectorGroup,
  dataStoreSpecsFromIds,
  getEngineDataStoreIds,
  listEngineDataStores,
} from './data-stores.js';

const cfg: GeminiClientConfig = {
  assistant: { project: 'p', location: 'global', engine: 'e', assistant: 'default_assistant' },
};
const tokens: TokenSource = { getAccessToken: async () => 'tok' };
const DS = 'projects/p/locations/global/collections/default_collection/dataStores';

function mockFetch(byUrl: (url: string) => unknown) {
  const fetchImpl = (async (input: RequestInfo | URL) => {
    const url = String(input);
    return new Response(JSON.stringify(byUrl(url)), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    });
  }) as unknown as typeof fetch;
  return fetchImpl;
}

describe('connectorGroup', () => {
  it('maps federated id prefixes to friendly groups', () => {
    expect(connectorGroup('msft-sharepoint-fed_123_file')).toBe('SharePoint');
    expect(connectorGroup('msft-onedrive-fed_123_file')).toBe('OneDrive');
    expect(connectorGroup('msft-outlook-fed_123_mail')).toBe('Outlook');
    expect(connectorGroup('acme-salesforce-fed_9_x')).toBe('Salesforce');
    expect(connectorGroup('weird-id')).toBe('Other');
  });
});

describe('getEngineDataStoreIds', () => {
  it('returns dataStoreIds from engines.get', async () => {
    const fetchImpl = mockFetch(() => ({
      name: 'projects/p/…/engines/e',
      dataStoreIds: ['msft-sharepoint-fed_1_file', 'msft-outlook-fed_1_mail'],
    }));
    const ids = await getEngineDataStoreIds(cfg, { tokens, fetchImpl });
    expect(ids).toEqual(['msft-sharepoint-fed_1_file', 'msft-outlook-fed_1_mail']);
  });
});

describe('listEngineDataStores', () => {
  it('enriches each id with displayName + connector + resourceName', async () => {
    const fetchImpl = mockFetch((url) => {
      if (url.endsWith('/engines/e'))
        return { dataStoreIds: ['msft-sharepoint-fed_1_file', 'msft-outlook-fed_1_mail'] };
      if (url.endsWith('/dataStores/msft-sharepoint-fed_1_file'))
        return { name: `${DS}/msft-sharepoint-fed_1_file`, displayName: 'file' };
      return { name: url.split('/').pop() }; // outlook one has no displayName → falls back to id
    });
    const stores = await listEngineDataStores(cfg, { tokens, fetchImpl });
    expect(stores).toEqual([
      {
        id: 'msft-sharepoint-fed_1_file',
        resourceName: `${DS}/msft-sharepoint-fed_1_file`,
        displayName: 'file',
        connector: 'SharePoint',
      },
      {
        id: 'msft-outlook-fed_1_mail',
        resourceName: `${DS}/msft-outlook-fed_1_mail`,
        displayName: 'msft-outlook-fed_1_mail',
        connector: 'Outlook',
      },
    ]);
  });
});

describe('dataStoreSpecsFromIds', () => {
  it('expands bare ids to resource names and passes full names through', () => {
    expect(dataStoreSpecsFromIds(cfg, ['msft-onedrive-fed_1_file', `${DS}/already-full`])).toEqual([
      { dataStore: `${DS}/msft-onedrive-fed_1_file` },
      { dataStore: `${DS}/already-full` },
    ]);
  });
});
