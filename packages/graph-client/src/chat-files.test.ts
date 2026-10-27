import { describe, expect, it, vi } from 'vitest';
import { GraphClient, encodeSharingUrl } from './graph-client.js';
import { isAllowedSharingUrl } from './config.js';

const tokens = { getGraphToken: vi.fn(async (_scopes: string[]) => 'tok') };

function fetchRouting(routes: Record<string, () => Response>) {
  return vi.fn(async (url: string | URL | Request) => {
    const key = Object.keys(routes).find((k) => String(url).includes(k));
    if (!key) return new Response('nope', { status: 404 });
    return routes[key]!();
  });
}

describe('encodeSharingUrl', () => {
  it('produces the u!<base64url> token Graph expects', () => {
    const token = encodeSharingUrl('https://contoso.sharepoint.com/:w:/s/team/abc?e=1');
    expect(token.startsWith('u!')).toBe(true);
    expect(token).not.toMatch(/[=+/]/);
  });
});

describe('isAllowedSharingUrl', () => {
  it('allows only https links on the tenant SharePoint hosts', () => {
    expect(isAllowedSharingUrl('https://contoso.sharepoint.com/:w:/s/x')).toBe(true);
    expect(isAllowedSharingUrl('https://contoso-my.sharepoint.com/personal/x')).toBe(true);
    expect(isAllowedSharingUrl('http://contoso.sharepoint.com/x')).toBe(false);
    expect(isAllowedSharingUrl('https://sharepoint.com.evil.example/x')).toBe(false);
    expect(isAllowedSharingUrl('https://contoso.sharepoint.us/x', ['sharepoint.us'])).toBe(true);
    expect(isAllowedSharingUrl('not a url')).toBe(false);
  });
});

const DECK = 'https://contoso.sharepoint.com/x/deck.pptx';
const SECRET = 'https://contoso-my.sharepoint.com/x/secret.docx';
const TRAP = 'https://evil.example/x/trap.pdf';

describe('GraphClient.listChatFiles', () => {
  it('resolves tenant-hosted reference attachments the user can open, and nothing else', async () => {
    const fetchImpl = fetchRouting({
      '/chats/19%3Aabc/messages': () =>
        new Response(
          JSON.stringify({
            value: [
              {
                id: 'm2',
                createdDateTime: '2026-10-27T01:00:00Z',
                from: { user: { displayName: 'Dana' } },
                attachments: [
                  { contentType: 'reference', contentUrl: DECK, name: 'deck.pptx' },
                  { contentType: 'reference', contentUrl: SECRET, name: 'secret.docx' },
                  { contentType: 'reference', contentUrl: TRAP, name: 'trap.pdf' },
                  { contentType: 'application/vnd.microsoft.card.adaptive', content: '{}' },
                ],
              },
              { id: 'm1', attachments: [] },
            ],
          }),
          { status: 200 },
        ),
      [`/shares/${encodeSharingUrl(DECK)}/driveItem`]: () =>
        new Response(
          JSON.stringify({
            id: 'item1',
            name: 'deck.pptx',
            size: 1234,
            webUrl: DECK,
            file: {
              mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            },
            parentReference: { driveId: 'drive1' },
          }),
          { status: 200 },
        ),
      [`/shares/${encodeSharingUrl(SECRET)}/driveItem`]: () =>
        new Response('forbidden', { status: 403 }),
      [`/shares/${encodeSharingUrl(TRAP)}/driveItem`]: () => {
        throw new Error('a foreign host must never be resolved');
      },
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    const files = await client.listChatFiles('19:abc');
    expect(files).toEqual([
      {
        id: 'item1',
        driveId: 'drive1',
        name: 'deck.pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        sizeBytes: 1234,
        webUrl: DECK,
        messageId: 'm2',
        sharedBy: 'Dana',
        sharedAt: '2026-10-27T01:00:00Z',
      },
    ]);
    // The chat listing asked for the chat scope; the share resolution asked for the files scopes.
    const scopes = tokens.getGraphToken.mock.calls.map((c) => c[0]);
    expect(scopes.some((s) => s.includes('Chat.Read'))).toBe(true);
    expect(scopes.some((s) => s.includes('Files.Read.All'))).toBe(true);
    expect(fetchImpl.mock.calls.some((c) => String(c[0]).includes(encodeSharingUrl(TRAP)))).toBe(
      false,
    );
  });
});

describe('GraphClient.getDriveItemContent', () => {
  it('downloads through the pre-authenticated download URL without the Graph bearer token', async () => {
    const calls: Array<{ url: string; auth: string | undefined }> = [];
    const fetchImpl = vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      const u = String(url);
      calls.push({
        url: u,
        auth: (init?.headers as Record<string, string> | undefined)?.Authorization,
      });
      if (u.includes('/drives/drive1/items/item1?'))
        return new Response(
          JSON.stringify({
            id: 'item1',
            size: 3,
            file: { mimeType: 'application/pdf' },
            '@microsoft.graph.downloadUrl': 'https://cdn.sharepoint.com/dl/item1?tempauth=x',
          }),
          { status: 200 },
        );
      if (u.startsWith('https://cdn.sharepoint.com/dl/item1'))
        return new Response(new Uint8Array([1, 2, 3]), {
          status: 200,
          headers: { 'content-type': 'application/pdf; charset=binary', 'content-length': '3' },
        });
      return new Response('nope', { status: 404 });
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    const got = await client.getDriveItemContent('item1', 'drive1');
    expect([...got.bytes]).toEqual([1, 2, 3]);
    expect(got.mimeType).toBe('application/pdf');
    expect(calls.find((c) => c.url.startsWith('https://cdn.sharepoint.com'))?.auth).toBeUndefined();
    expect(calls.find((c) => c.url.includes('/drives/'))?.auth).toBe('Bearer tok');
  });

  it('falls back to /content when the item carries no download URL', async () => {
    const fetchImpl = fetchRouting({
      '/me/drive/items/plain?': () =>
        new Response(JSON.stringify({ id: 'plain', size: 2 }), { status: 200 }),
      '/me/drive/items/plain/content': () => new Response(new Uint8Array([7, 8]), { status: 200 }),
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    const got = await client.getDriveItemContent('plain');
    expect([...got.bytes]).toEqual([7, 8]);
  });

  it('refuses a file over the caller byte limit before downloading', async () => {
    const fetchImpl = fetchRouting({
      '/me/drive/items/big?': () =>
        new Response(JSON.stringify({ id: 'big', size: 10 }), { status: 200 }),
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    await expect(client.getDriveItemContent('big', undefined, { maxBytes: 4 })).rejects.toThrow(
      'read limit',
    );
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });
});
