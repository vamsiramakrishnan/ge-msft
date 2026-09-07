import { describe, expect, it, vi } from 'vitest';
import type { AssistRequest, SseEvent } from '@ge/contracts';
import type { GeminiClientConfig } from './config.js';
import {
  buildStreamAssistRequest,
  StreamAssistClient,
  type StreamOptions,
} from './stream-assist.js';

const config: GeminiClientConfig = {
  assistant: { project: 'project', location: 'eu', engine: 'engine' },
  identity: 'user@example.com',
};
const request: AssistRequest = {
  intent: 'ask',
  query: 'Summarize this input.',
  unit: { connectors: [], surfaceContext: { kind: 'word', selection: 'Explicit input' } },
};
const responseFrames = [
  {
    sessionInfo: { session: 'projects/project/locations/eu/sessions/unexpected-session' },
    answer: {
      state: 'SUCCEEDED',
      replies: [{ groundedContent: { content: { text: 'Summary' } } }],
    },
  },
];

function setup(cfg = config) {
  const tokens = { getAccessToken: vi.fn(async () => 'test-token') };
  const sent: { url: string; body: Record<string, unknown> }[] = [];
  const fetcher: typeof fetch = vi.fn(async (url, init) => {
    sent.push({
      url: String(url),
      body: JSON.parse(String(init?.body)) as Record<string, unknown>,
    });
    return new Response(JSON.stringify(responseFrames), { status: 200 });
  });
  return { client: new StreamAssistClient(tokens, cfg, fetcher), tokens, fetcher, sent };
}

async function collect(events: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const event of events) out.push(event);
  return out;
}

describe('StreamAssistClient explicit sessionless requests', () => {
  it.each([undefined, '', '-'])(
    'sends isSessionLess=true without a session for %j',
    async (session) => {
      const { client, sent } = setup();
      const events = await collect(client.stream(request, { isSessionLess: true, session }));
      expect(sent).toHaveLength(1);
      expect(sent[0]!.url).toContain('/v1alpha/projects/');
      expect(sent[0]!.url).toContain(':streamAssist');
      expect(sent[0]!.body).toMatchObject({ isSessionLess: true });
      expect(sent[0]!.body).not.toHaveProperty('session');
      expect(events).toContainEqual({ type: 'token', text: 'Summary' });
      const provenance = events.find((event) => event.type === 'provenance');
      expect(provenance).toBeDefined();
      expect(provenance).not.toHaveProperty('payload.sessionId');
      expect(events.at(-1)).toEqual({ type: 'done' });
    },
  );

  it.each([undefined, false])(
    'retains the existing stateful default for %j',
    async (isSessionLess) => {
      const { client, sent } = setup();
      const events = await collect(
        client.stream(request, { session: 'existing-session', isSessionLess }),
      );
      expect(sent[0]!.body).toHaveProperty('session', 'existing-session');
      expect(sent[0]!.body).not.toHaveProperty('isSessionLess');
      expect(events.find((event) => event.type === 'provenance')).toHaveProperty(
        'payload.sessionId',
        'projects/project/locations/eu/sessions/unexpected-session',
      );
    },
  );

  it.each(['existing-session', 'projects/project/locations/eu/sessions/123', ' ', ' - '])(
    'rejects incompatible session %j before acquiring a token or fetching',
    async (session) => {
      const { client, tokens, fetcher } = setup();
      const events = await collect(client.stream(request, { isSessionLess: true, session }));
      expect(events).toEqual([
        {
          type: 'error',
          code: 'invalid_request',
          message: expect.stringContaining('cannot resume'),
        },
      ]);
      expect(tokens.getAccessToken).not.toHaveBeenCalled();
      expect(fetcher).not.toHaveBeenCalled();
    },
  );

  it('rejects session-bound file ids without silently discarding grounding', async () => {
    const { client, tokens, fetcher } = setup();
    const events = await collect(
      client.stream(request, {
        isSessionLess: true,
        grounding: { fileIds: ['session-file'] },
      }),
    );
    expect(events).toEqual([
      {
        type: 'error',
        code: 'invalid_request',
        message: expect.stringContaining('session context files'),
      },
    ]);
    expect(tokens.getAccessToken).not.toHaveBeenCalled();
    expect(fetcher).not.toHaveBeenCalled();
  });

  it('preserves explicit inline, indexed-document and datastore grounding', async () => {
    const { client, sent } = setup();
    await collect(
      client.stream(request, {
        isSessionLess: true,
        grounding: {
          queryParts: [
            { text: 'Bounded execution capsule' },
            { documentReference: { documentName: 'indexed-document' } },
          ],
          dataStoreSpecs: [{ dataStore: 'selected-store', filter: 'owner:ANY("user")' }],
          fileIds: [],
        },
      }),
    );
    expect(sent[0]!.body).toMatchObject({
      isSessionLess: true,
      query: {
        parts: [
          { text: 'Bounded execution capsule' },
          { documentReference: { documentName: 'indexed-document' } },
          { text: request.query },
        ],
      },
      toolsSpec: {
        vertexAiSearchSpec: {
          dataStoreSpecs: [{ dataStore: 'selected-store', filter: 'owner:ANY("user")' }],
        },
      },
    });
  });

  it('allows widget catalog metadata while continuing to use the direct endpoint', async () => {
    const { client, sent } = setup({ ...config, widget: { configId: 'catalog-widget' } });
    await collect(client.stream(request, { isSessionLess: true }));
    expect(sent[0]!.body).toHaveProperty('isSessionLess', true);
    expect(sent[0]!.url).toContain(':streamAssist');
    expect(sent[0]!.url).not.toContain('widgetStreamAssist');
  });

  it('validates the public request builder as well as the streaming client', () => {
    expect(() =>
      buildStreamAssistRequest(request, config, 'existing', undefined, 'default', undefined, {
        isSessionLess: true,
      }),
    ).toThrow('cannot resume');
    expect(() =>
      buildStreamAssistRequest(
        request,
        config,
        undefined,
        undefined,
        'default',
        { fileIds: ['file'] },
        { isSessionLess: true },
      ),
    ).toThrow('session context files');
    expect(buildStreamAssistRequest(request, config)).not.toHaveProperty('isSessionLess');
  });

  it('snapshots sessionless mode so mutable options cannot leak response session provenance', async () => {
    const options: StreamOptions = { isSessionLess: true, session: '-' };
    const fetcher: typeof fetch = vi.fn(async () => {
      options.isSessionLess = false;
      options.session = 'stale-session';
      return new Response(JSON.stringify(responseFrames));
    });
    const client = new StreamAssistClient(
      { getAccessToken: async () => 'test-token' },
      config,
      fetcher,
    );
    const events = await collect(client.stream(request, options));
    expect(events.find((event) => event.type === 'provenance')).not.toHaveProperty(
      'payload.sessionId',
    );
  });
});
