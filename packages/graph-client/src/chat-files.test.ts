import { describe, expect, it, vi } from 'vitest';
import { GraphClient, encodeSharingUrl } from './graph-client.js';

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

describe('GraphClient.listChatFiles', () => {
  it('resolves reference attachments to readable drive items, skipping what the user cannot open', async () => {
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
                  {
                    contentType: 'reference',
                    contentUrl: 'https://x/deck.pptx',
                    name: 'deck.pptx',
                  },
                  {
                    contentType: 'reference',
                    contentUrl: 'https://x/secret.docx',
                    name: 'secret.docx',
                  },
                  { contentType: 'application/vnd.microsoft.card.adaptive', content: '{}' },
                ],
              },
              { id: 'm1', attachments: [] },
            ],
          }),
          { status: 200 },
        ),
      [`/shares/${encodeSharingUrl('https://x/deck.pptx')}/driveItem`]: () =>
        new Response(
          JSON.stringify({
            id: 'item1',
            name: 'deck.pptx',
            size: 1234,
            webUrl: 'https://x/deck.pptx',
            file: {
              mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            },
            parentReference: { driveId: 'drive1' },
          }),
          { status: 200 },
        ),
      [`/shares/${encodeSharingUrl('https://x/secret.docx')}/driveItem`]: () =>
        new Response('forbidden', { status: 403 }),
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
        webUrl: 'https://x/deck.pptx',
        messageId: 'm2',
        sharedBy: 'Dana',
        sharedAt: '2026-10-27T01:00:00Z',
      },
    ]);
    // The chat listing asked for the chat scope; the share resolution asked for the files scopes.
    const scopes = tokens.getGraphToken.mock.calls.map((c) => c[0]);
    expect(scopes.some((s) => s.includes('Chat.Read'))).toBe(true);
    expect(scopes.some((s) => s.includes('Files.Read.All'))).toBe(true);
  });
});

describe('GraphClient.getDriveItemContent', () => {
  it('downloads bytes as the user and reports the content type', async () => {
    const fetchImpl = fetchRouting({
      '/drives/drive1/items/item1/content': () =>
        new Response(new Uint8Array([1, 2, 3]), {
          status: 200,
          headers: { 'content-type': 'application/pdf; charset=binary', 'content-length': '3' },
        }),
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    const got = await client.getDriveItemContent('item1', 'drive1');
    expect([...got.bytes]).toEqual([1, 2, 3]);
    expect(got.mimeType).toBe('application/pdf');
  });

  it('refuses a file over the caller byte limit', async () => {
    const fetchImpl = fetchRouting({
      '/me/drive/items/big/content': () =>
        new Response(new Uint8Array(10), { status: 200, headers: { 'content-length': '10' } }),
    });
    const client = new GraphClient(tokens, {}, fetchImpl as never);
    await expect(client.getDriveItemContent('big', undefined, { maxBytes: 4 })).rejects.toThrow(
      'read limit',
    );
  });
});
