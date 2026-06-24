import { describe, it, expect, vi } from 'vitest';
import { ResolvedContextSchema, EstateRefSchema } from '@ge/contracts';
import {
  GraphClient,
  userToContext,
  driveItemToContext,
  eventToContext,
  messageToContext,
  type GraphTokenSource,
} from './graph-client.js';
import { GRAPH_SCOPES, GRAPH_BASE_URL, graphUrl } from './config.js';

function jsonResponse(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/** A token source whose token and invalidation are observable. */
function fakeTokens(
  token = 'graph-token',
): GraphTokenSource & { invalidate: ReturnType<typeof vi.fn> } {
  return { getGraphToken: () => Promise.resolve(token), invalidate: vi.fn() };
}

type AnyMock = { mock: { calls: unknown[][] } };
function lastInit(f: AnyMock, call = 0): RequestInit {
  return f.mock.calls[call]![1] as RequestInit;
}
function lastUrl(f: AnyMock, call = 0): string {
  return f.mock.calls[call]![0] as string;
}

// ---------------------------------------------------------------------------
// config: graphUrl branch coverage (absolute pass-through, base override, slash join)
// ---------------------------------------------------------------------------

describe('graphUrl', () => {
  it('joins a relative path onto the default base', () => {
    expect(graphUrl({}, '/me/messages/x')).toBe(`${GRAPH_BASE_URL}/me/messages/x`);
  });

  it('inserts a slash when the path lacks a leading one', () => {
    expect(graphUrl({}, 'me/drive')).toBe(`${GRAPH_BASE_URL}/me/drive`);
  });

  it('passes an absolute http url straight through (nextLink case)', () => {
    const abs = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=abc';
    expect(graphUrl({}, abs)).toBe(abs);
  });

  it('honours a sovereign-cloud baseUrl override and trims its trailing slash', () => {
    const cfg = { baseUrl: 'https://graph.microsoft.us/v1.0/' };
    expect(graphUrl(cfg, '/me')).toBe('https://graph.microsoft.us/v1.0/me');
  });
});

// ---------------------------------------------------------------------------
// typed reads: each verb hits the right delegated-scope path with a Bearer token
// ---------------------------------------------------------------------------

describe('GraphClient typed reads — delegated scopes & paths', () => {
  it('getEvent reads /me/events with the calendar scope', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'e1', subject: 'Sync' }));
    const tokens = fakeTokens();
    const spy = vi.spyOn(tokens, 'getGraphToken');
    const client = new GraphClient(tokens, {}, f as never);

    const ev = await client.getEvent('e1');

    expect(ev.subject).toBe('Sync');
    expect(lastUrl(f)).toContain('/me/events/e1');
    expect(lastInit(f).method).toBe('GET');
    expect(spy).toHaveBeenCalledWith([...GRAPH_SCOPES.calendar]);
  });

  it('getUser reads /users/{id} with the people scope', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'u1', displayName: 'Vamsi' }));
    const tokens = fakeTokens();
    const spy = vi.spyOn(tokens, 'getGraphToken');
    const client = new GraphClient(tokens, {}, f as never);

    await client.getUser('u1');

    expect(lastUrl(f)).toContain('/users/u1');
    expect(spy).toHaveBeenCalledWith([...GRAPH_SCOPES.people]);
  });

  it('getDriveItem WITHOUT a driveId targets the personal /me/drive path', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'd1', name: 'Notes.docx' }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.getDriveItem('d1');

    expect(lastUrl(f)).toContain('/me/drive/items/d1');
    expect(lastUrl(f)).not.toContain('/drives/');
  });

  it('getDriveItem WITH a driveId targets the explicit /drives/{driveId} path', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'd1', name: 'Shared.docx' }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.getDriveItem('d1', 'drvX');

    expect(lastUrl(f)).toContain('/drives/drvX/items/d1');
  });

  it('url-encodes ids so a path-traversal id cannot escape the resource path', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'm1' }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.getMessage('a/b c?$x');

    const url = lastUrl(f);
    expect(url).toContain('/me/messages/a%2Fb%20c%3F%24x');
    expect(url).not.toContain('/me/messages/a/b');
  });

  it('does NOT attach a Content-Type header on GET reads (no body)', async () => {
    const f = vi.fn(async () => jsonResponse({ id: 'm1' }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.getMessage('m1');

    const headers = lastInit(f).headers as Record<string, string>;
    expect(headers['Content-Type']).toBeUndefined();
    expect(headers.Authorization).toBe('Bearer graph-token');
    expect(lastInit(f).body).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// Microsoft Search transport: POST shape, scopes, content-type, entityTypes & size
// ---------------------------------------------------------------------------

describe('GraphClient.search — POST transport', () => {
  it('POSTs to /search/query with a JSON body and Content-Type header', async () => {
    const f = vi.fn(async () => jsonResponse({ value: [] }));
    const tokens = fakeTokens();
    const spy = vi.spyOn(tokens, 'getGraphToken');
    const client = new GraphClient(tokens, {}, f as never);

    await client.search('vendor risk');

    const init = lastInit(f);
    expect(init.method).toBe('POST');
    expect(lastUrl(f)).toContain('/search/query');
    expect((init.headers as Record<string, string>)['Content-Type']).toBe('application/json');
    expect(spy).toHaveBeenCalledWith([...GRAPH_SCOPES.search]);
  });

  it('serialises the queryString, entityTypes, and size into the request body', async () => {
    const f = vi.fn(async () => jsonResponse({ value: [] }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.search('quarterly numbers', ['message', 'event'], 25);

    const body = JSON.parse(lastInit(f).body as string);
    expect(body.requests).toHaveLength(1);
    expect(body.requests[0]).toMatchObject({
      entityTypes: ['message', 'event'],
      query: { queryString: 'quarterly numbers' },
      from: 0,
      size: 25,
    });
  });

  it('defaults to driveItem+listItem entity types and size 10', async () => {
    const f = vi.fn(async () => jsonResponse({ value: [] }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await client.search('anything');

    const body = JSON.parse(lastInit(f).body as string);
    expect(body.requests[0].entityTypes).toEqual(['driveItem', 'listItem']);
    expect(body.requests[0].size).toBe(10);
  });

  it('returns an empty list when the response has no value array', async () => {
    const f = vi.fn(async () => jsonResponse({}));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const refs = await client.search('nothing');
    expect(refs).toEqual([]);
  });

  it('flattens hits across multiple containers and skips hits with no resource', async () => {
    const f = vi.fn(async () =>
      jsonResponse({
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  { hitId: 'noRes', summary: 'dropped — no resource' },
                  {
                    hitId: 'h1',
                    resource: {
                      '@odata.type': '#microsoft.graph.message',
                      id: 'm1',
                      subject: 'Re: SLA',
                    },
                  },
                ],
              },
              {
                hits: [
                  {
                    hitId: 'h2',
                    resource: {
                      '@odata.type': '#microsoft.graph.event',
                      id: 'e1',
                      subject: 'Review',
                    },
                  },
                ],
              },
            ],
          },
        ],
      }),
    );
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const refs = await client.search('mix', ['message', 'event']);

    expect(refs).toHaveLength(2);
    for (const r of refs) expect(() => EstateRefSchema.parse(r)).not.toThrow();
    expect(refs[0]).toMatchObject({ source: 'mail', id: 'm1', title: 'Re: SLA' });
    expect(refs[1]).toMatchObject({ source: 'calendar', id: 'e1', title: 'Review' });
  });
});

// ---------------------------------------------------------------------------
// hitToEstateRef branch matrix (exercised through the public .search surface)
// ---------------------------------------------------------------------------

describe('GraphClient.search — hit → EstateRef mapping branches', () => {
  function searchReturning(
    resource: Record<string, unknown> | undefined,
    summary?: string,
    hitId?: string,
  ) {
    const f = vi.fn(async () =>
      jsonResponse({ value: [{ hitsContainers: [{ hits: [{ hitId, summary, resource }] }] }] }),
    );
    return { f, client: new GraphClient(fakeTokens(), {}, f as never) };
  }

  it('maps a listItem odata type to a site source', async () => {
    const { client } = searchReturning({
      '@odata.type': '#microsoft.graph.listItem',
      id: 'li1',
      displayName: 'Policy row',
    });
    const refs = await client.search('x', ['listItem']);
    expect(refs[0]).toMatchObject({ source: 'site', id: 'li1', title: 'Policy row' });
  });

  it('falls back to drive-item for an unrecognised odata type', async () => {
    const { client } = searchReturning({
      '@odata.type': '#microsoft.graph.somethingNew',
      id: 'q1',
    });
    const refs = await client.search('x');
    expect(refs[0]).toMatchObject({ source: 'drive-item', id: 'q1' });
  });

  it('uses hitId as the id when the resource carries none', async () => {
    const { client } = searchReturning(
      { '@odata.type': '#microsoft.graph.message', subject: 'No id here' },
      undefined,
      'fallback-hit',
    );
    const refs = await client.search('x', ['message']);
    expect(refs[0]).toMatchObject({ source: 'mail', id: 'fallback-hit', title: 'No id here' });
  });

  it('drops a hit whose resource has neither id nor hitId', async () => {
    const { client } = searchReturning({
      '@odata.type': '#microsoft.graph.message',
      subject: 'orphan',
    });
    const refs = await client.search('x', ['message']);
    expect(refs).toEqual([]);
  });

  it('carries webLink (not just webUrl) and the summary preview onto the ref', async () => {
    const { client } = searchReturning(
      {
        '@odata.type': '#microsoft.graph.message',
        id: 'm2',
        subject: 'Has weblink',
        webLink: 'https://outlook/m2',
      },
      'preview text',
    );
    const refs = await client.search('x', ['message']);
    expect(refs[0]).toMatchObject({
      source: 'mail',
      webUrl: 'https://outlook/m2',
      preview: 'preview text',
    });
  });

  it('extracts the driveId from a driveItem hit parentReference', async () => {
    const { client } = searchReturning({
      '@odata.type': '#microsoft.graph.driveItem',
      id: 'd5',
      name: 'Deck.pptx',
      parentReference: { driveId: 'drvP' },
    });
    const refs = await client.search('x');
    expect(refs[0]).toMatchObject({
      source: 'drive-item',
      id: 'd5',
      driveId: 'drvP',
      title: 'Deck.pptx',
    });
  });

  it('omits driveId for a driveItem hit without a parentReference', async () => {
    const { client } = searchReturning({
      '@odata.type': '#microsoft.graph.driveItem',
      id: 'd6',
      name: 'Loose.docx',
    });
    const refs = await client.search('x');
    expect(refs[0]).toMatchObject({ source: 'drive-item', id: 'd6' });
    expect((refs[0] as { driveId?: string }).driveId).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------
// resolveEstateRef: the calendar, site, and person switch arms
// ---------------------------------------------------------------------------

describe('GraphClient.resolveEstateRef — source arms', () => {
  it('resolves a calendar ref by fetching the event and mapping it to context', async () => {
    const f = vi.fn(async () =>
      jsonResponse({
        id: 'e1',
        subject: 'Vendor review',
        start: { dateTime: '2026-06-22T09:00:00' },
      }),
    );
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const ctx = await client.resolveEstateRef({ source: 'calendar', id: 'e1' });

    expect(lastUrl(f)).toContain('/me/events/e1');
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect((ctx[0]!.value as { text: string }).text).toContain('Event: Vendor review');
  });

  it('resolves a site ref WITHOUT any network call (titled reference back to source)', async () => {
    const f = vi.fn(async () => jsonResponse({}));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const ctx = await client.resolveEstateRef({
      source: 'site',
      id: 'li9',
      title: 'Q3 list item',
      webUrl: 'https://sp/list/li9',
    });

    expect(f).not.toHaveBeenCalled();
    expect(ctx[0]!.value).toMatchObject({
      as: 'indexed-document',
      documentName: 'li9',
      title: 'Q3 list item',
    });
    expect(ctx[0]!.ref).toMatchObject({ kind: 'indexed-document', preview: 'https://sp/list/li9' });
  });

  it('resolves a site ref with NO title/webUrl to the generic "Item" reference', async () => {
    const client = new GraphClient(fakeTokens(), {}, vi.fn() as never);
    const ctx = await client.resolveEstateRef({ source: 'site', id: 'liBare' });
    expect(ctx[0]!.ref).toMatchObject({
      id: 'graph:site:liBare',
      kind: 'indexed-document',
      title: 'Item',
    });
    expect((ctx[0]!.ref as { preview?: string }).preview).toBeUndefined();
    expect(ctx[0]!.value).toMatchObject({
      as: 'indexed-document',
      documentName: 'liBare',
      title: 'Item',
    });
  });

  it('resolves a mail-thread ref through the message read path', async () => {
    const f = vi.fn(async () =>
      jsonResponse({ id: 't1', subject: 'Thread', body: { content: 'hi' } }),
    );
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const ctx = await client.resolveEstateRef({ source: 'mail-thread', id: 't1' });

    expect(lastUrl(f)).toContain('/me/messages/t1');
    expect((ctx[0]!.value as { text: string }).text).toContain('Subject: Thread');
  });
});

// ---------------------------------------------------------------------------
// 401 refresh / error-path semantics
// ---------------------------------------------------------------------------

describe('GraphClient — auth refresh & failure handling', () => {
  it('does NOT retry on 401 when the token source cannot invalidate', async () => {
    const f = vi.fn(async () => new Response('expired', { status: 401 }));
    // token source without invalidate()
    const client = new GraphClient({ getGraphToken: () => Promise.resolve('t') }, {}, f as never);

    await expect(client.getMessage('m1')).rejects.toThrow(/failed \(401\)/);
    expect(f).toHaveBeenCalledOnce();
  });

  it('surfaces a still-401 after a refresh attempt as a thrown error', async () => {
    const inval = vi.fn();
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(new Response('still expired', { status: 401 }));
    const client = new GraphClient(
      { getGraphToken: () => Promise.resolve('t'), invalidate: inval },
      {},
      f as never,
    );

    await expect(client.getUser('u1')).rejects.toThrow(/Graph GET \/users\/u1 failed \(401\)/);
    expect(inval).toHaveBeenCalledOnce();
    expect(f).toHaveBeenCalledTimes(2);
  });

  it('throws on a 429 throttling response (no infinite retry) and includes the body snippet', async () => {
    const f = vi.fn(async () => new Response('rate limited, retry-after 30', { status: 429 }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await expect(client.search('busy')).rejects.toThrow(
      /Graph POST \/search\/query failed \(429\): rate limited, retry-after 30/,
    );
    expect(f).toHaveBeenCalledOnce();
  });

  it('throws on a 500 server error with the method, path and status', async () => {
    const f = vi.fn(async () => new Response('boom', { status: 500 }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await expect(client.getEvent('e1')).rejects.toThrow(
      /Graph GET \/me\/events\/e1 failed \(500\)/,
    );
  });

  it('truncates a very long error body to 300 chars in the thrown message', async () => {
    const huge = 'x'.repeat(1000);
    const f = vi.fn(async () => new Response(huge, { status: 400 }));
    const client = new GraphClient(fakeTokens(), {}, f as never);

    const err = await client.getUser('u1').catch((e: Error) => e);
    const msg = (err as Error).message;
    // 'x' run in the message should be capped at 300
    const run = msg.match(/x+/)?.[0] ?? '';
    expect(run.length).toBe(300);
  });

  it('still throws (without crashing) when the error body cannot be read as text', async () => {
    const broken = {
      ok: false,
      status: 502,
      text: () => Promise.reject(new Error('stream broke')),
    } as unknown as Response;
    const f = vi.fn(async () => broken);
    const client = new GraphClient(fakeTokens(), {}, f as never);

    await expect(client.getMessage('m1')).rejects.toThrow(/failed \(502\)/);
  });
});

// ---------------------------------------------------------------------------
// mappers not previously covered: userToContext fallbacks & driveItemToContext fallbacks
// ---------------------------------------------------------------------------

describe('userToContext', () => {
  it('builds person context from displayName, jobTitle and mail', () => {
    const ctx = userToContext({ id: 'u1', displayName: 'Vamsi', jobTitle: 'PM', mail: 'v@acme' });
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const text = (ctx[0]!.value as { text: string }).text;
    expect(text).toContain('Name: Vamsi');
    expect(text).toContain('Title: PM');
    expect(text).toContain('Email: v@acme');
    // chunk ids derive from the source id (graph:person:<id>) plus a chunk suffix
    expect(ctx[0]!.ref.id).toMatch(/^graph:person:u1/);
  });

  it('falls back to userPrincipalName when mail is absent', () => {
    const withUpn = userToContext({ id: 'u2', userPrincipalName: 'upn@acme' });
    expect((withUpn[0]!.value as { text: string }).text).toContain('Email: upn@acme');
  });

  it('produces no context chunks for a bare user with no displayable fields (empty body)', () => {
    // Every field is empty → the joined text is empty → no content chunks are emitted.
    expect(userToContext({})).toEqual([]);
  });
});

describe('eventToContext — optional-field branches', () => {
  it('omits End and Location lines when absent, and uses bodyPreview when body content is missing', () => {
    const ctx = eventToContext({
      id: 'e2',
      subject: 'Standup',
      start: { dateTime: '2026-06-23T09:00:00' },
      bodyPreview: 'daily',
    });
    const text = (ctx[0]!.value as { text: string }).text;
    expect(text).toContain('Event: Standup');
    expect(text).toContain('Start: 2026-06-23T09:00:00');
    expect(text).not.toContain('End:');
    expect(text).not.toContain('Location:');
    expect(text).toContain('daily');
  });

  it('routes an HTML event body through the html→markdown reduction (tags stripped)', () => {
    const ctx = eventToContext({
      id: 'e3',
      subject: 'Brief',
      body: { contentType: 'HTML', content: '<p>hello <strong>world</strong></p>' },
    });
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const text = (ctx[0]!.value as { text: string }).text;
    // contentType:'HTML' must drive format:'html' so the markup is reduced, not passed through.
    expect(text).toContain('Event: Brief');
    expect(text).toContain('hello **world**');
    expect(text).not.toContain('<p>');
    expect(text).not.toContain('<strong>');
  });
});

describe('messageToContext — optional-field branches', () => {
  it('uses bodyPreview when the message body is absent and labels it as an email reference', () => {
    const ctx = messageToContext({ id: 'm9', subject: 'No body', bodyPreview: 'preview only' });
    const text = (ctx[0]!.value as { text: string }).text;
    expect(text).toContain('Subject: No body');
    expect(text).toContain('preview only');
    expect(ctx[0]!.ref.surface).toBe('outlook');
  });

  it('renders a sender with no name as an empty-name angle-bracket address (untrusted shape tolerated)', () => {
    const ctx = messageToContext({
      id: 'm10',
      from: { emailAddress: { address: 'x@y' } },
      body: { content: 'b' },
    });
    expect((ctx[0]!.value as { text: string }).text).toContain('From:  <x@y>');
  });

  it('falls back to the "Email" title when the message has no subject', () => {
    const ctx = messageToContext({ id: 'm11', body: { content: 'no subject body' } });
    // with no subject, the source title falls back to the generic 'Email' label
    expect(ctx[0]!.ref.title).toBe('Email');
    expect((ctx[0]!.value as { text: string }).text).toContain('no subject body');
  });
});

describe('driveItemToContext', () => {
  it('prefers the item name, item driveId, and webUrl preview', () => {
    const ctx = driveItemToContext(
      {
        id: 'd1',
        name: 'MSA.docx',
        webUrl: 'https://sp/MSA',
        size: 4096,
        parentReference: { driveId: 'drvItem' },
      },
      { source: 'drive-item', id: 'd1', driveId: 'drvRef', title: 'Ref title' },
    );
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx[0]!.value).toMatchObject({
      as: 'drive-document',
      driveId: 'drvItem',
      documentName: 'd1',
      title: 'MSA.docx',
    });
    expect(ctx[0]!.ref).toMatchObject({ preview: 'https://sp/MSA', sizeBytes: 4096 });
  });

  it('falls back to the ref title/id/driveId when the drive item omits them', () => {
    const ctx = driveItemToContext(
      {},
      { source: 'drive-item', id: 'dref', driveId: 'drvRef', title: 'Ref title' },
    );
    expect(ctx[0]!.value).toMatchObject({
      driveId: 'drvRef',
      documentName: 'dref',
      title: 'Ref title',
    });
    // no webUrl/size on item → no preview/sizeBytes keys
    expect((ctx[0]!.ref as { preview?: string }).preview).toBeUndefined();
    expect((ctx[0]!.ref as { sizeBytes?: number }).sizeBytes).toBeUndefined();
  });

  it('uses the generic "File" title when neither item nor ref names it', () => {
    const ctx = driveItemToContext({ id: 'dx' }, { source: 'drive-item', id: 'dx' });
    expect(ctx[0]!.value).toMatchObject({ title: 'File', driveId: '' });
  });
});
