import { describe, it, expect, vi } from 'vitest';
import { ResolvedContextSchema, EstateRefSchema } from '@ge/contracts';
import { GraphClient, messageToContext, eventToContext } from './graph-client.js';

function jsonResponse(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

const tokens = { getGraphToken: () => Promise.resolve('graph-token'), invalidate: vi.fn() };

describe('graph mappers (pure)', () => {
  it('maps a message to attach-ready, labelled context', () => {
    const ctx = messageToContext({
      id: 'm1',
      subject: 'SLA concerns',
      from: { emailAddress: { name: 'Vendor', address: 'v@acme' } },
      receivedDateTime: '2026-06-21T10:00:00Z',
      body: { contentType: 'text', content: 'The SLA is 99.5%.' },
    });
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const text = (ctx[0]!.value as { text: string }).text;
    expect(text).toContain('Subject: SLA concerns');
    expect(text).toContain('From: Vendor <v@acme>');
    expect(text).toContain('99.5%');
    expect(ctx[0]!.ref.surface).toBe('outlook');
  });

  it('maps a calendar event', () => {
    const ctx = eventToContext({
      id: 'e1',
      subject: 'Vendor review',
      start: { dateTime: '2026-06-22T09:00:00' },
      location: { displayName: 'Room 4' },
      bodyPreview: 'Quarterly check-in',
    });
    const text = (ctx[0]!.value as { text: string }).text;
    expect(text).toContain('Event: Vendor review');
    expect(text).toContain('Location: Room 4');
  });
});

describe('GraphClient.resolveEstateRef', () => {
  it('reads a mail item via Graph and resolves it to context', async () => {
    const f = vi.fn(async (url: string) => {
      expect(url).toContain('/me/messages/m1');
      return jsonResponse({
        id: 'm1',
        subject: 'Risk',
        from: { emailAddress: { address: 'a@b' } },
        body: { contentType: 'html', content: '<p>hi</p>' },
      });
    });
    const client = new GraphClient(tokens, {}, f as never);
    const ctx = await client.resolveEstateRef({ source: 'mail', id: 'm1' });
    expect(f).toHaveBeenCalledOnce();
    const init = (f.mock.calls[0] as unknown[])[1] as RequestInit;
    expect(init.headers).toMatchObject({ Authorization: 'Bearer graph-token' });
    expect(ctx[0]!.value.as).toBe('text');
  });

  it('resolves a drive-item to a reference (reference-over-inline)', async () => {
    const f = vi.fn(async () =>
      jsonResponse({
        id: 'd1',
        name: 'MSA.docx',
        webUrl: 'https://sp/MSA',
        parentReference: { driveId: 'drv1' },
      }),
    );
    const client = new GraphClient(tokens, {}, f as never);
    const ctx = await client.resolveEstateRef({ source: 'drive-item', id: 'd1', driveId: 'drv1' });
    expect(ctx[0]!.value).toMatchObject({
      as: 'drive-document',
      driveId: 'drv1',
      title: 'MSA.docx',
    });
  });
});

describe('GraphClient.search', () => {
  it('maps Microsoft Search hits to EstateRefs', async () => {
    const f = vi.fn(async (url: string) => {
      expect(url).toContain('/search/query');
      return jsonResponse({
        value: [
          {
            hitsContainers: [
              {
                hits: [
                  {
                    hitId: 'h1',
                    summary: 'vendor risk policy…',
                    resource: {
                      '@odata.type': '#microsoft.graph.driveItem',
                      id: 'd9',
                      name: 'Policy.pdf',
                      webUrl: 'https://sp/policy',
                      parentReference: { driveId: 'drvX' },
                    },
                  },
                ],
              },
            ],
          },
        ],
      });
    });
    const client = new GraphClient(tokens, {}, f as never);
    const refs = await client.search('vendor risk', ['driveItem']);
    expect(refs).toHaveLength(1);
    for (const r of refs) expect(() => EstateRefSchema.parse(r)).not.toThrow();
    expect(refs[0]).toMatchObject({
      source: 'drive-item',
      id: 'd9',
      driveId: 'drvX',
      title: 'Policy.pdf',
    });
  });
});

describe('GraphClient auth', () => {
  it('retries once on 401 after invalidating the token', async () => {
    const inval = vi.fn();
    const f = vi
      .fn()
      .mockResolvedValueOnce(new Response('expired', { status: 401 }))
      .mockResolvedValueOnce(jsonResponse({ id: 'u1', displayName: 'Vamsi', mail: 'v@acme' }));
    const client = new GraphClient(
      { getGraphToken: () => Promise.resolve('t'), invalidate: inval },
      {},
      f as never,
    );
    const ctx = await client.resolveEstateRef({ source: 'person', id: 'u1' });
    expect(inval).toHaveBeenCalledOnce();
    expect(f).toHaveBeenCalledTimes(2);
    expect((ctx[0]!.value as { text: string }).text).toContain('Name: Vamsi');
  });

  it('throws a helpful error on a non-401 Graph failure', async () => {
    const f = vi.fn(async () => new Response('forbidden', { status: 403 }));
    const client = new GraphClient(tokens, {}, f as never);
    await expect(client.getMessage('x')).rejects.toThrow(
      /Graph GET \/me\/messages\/x failed \(403\)/,
    );
  });
});
