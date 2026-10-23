import { afterEach, describe, expect, it, vi } from 'vitest';
import { liveFetch, type TranscriptEntry } from './live-env.js';

afterEach(() => {
  vi.unstubAllGlobals();
});

const url = 'https://discoveryengine.googleapis.com/v1alpha/projects/test:streamAssist';
const flush = async () => {
  for (let i = 0; i < 8; i++) await Promise.resolve();
};

describe('live scenario transport', () => {
  it('delivers the response and first bytes before the recorder sees EOF', async () => {
    let source!: ReadableStreamDefaultController<Uint8Array>;
    const body = new ReadableStream<Uint8Array>({ start: (controller) => (source = controller) });
    const chunk = new TextEncoder().encode(
      '[{"answer":{"replies":[{"groundedContent":{"content":{"text":"Hello"}}}]}}]',
    );
    source.enqueue(chunk);
    const fetchMock = vi.fn(
      async (_input: RequestInfo | URL, _init?: RequestInit) => new Response(body, { status: 200 }),
    );
    vi.stubGlobal('fetch', fetchMock);
    const transcript: TranscriptEntry[] = [];
    const pending = liveFetch({ VITE_WIF_USER_PROJECT: 'quota-project' }, transcript)(url, {
      body: '{"query":"hello"}',
    });
    let response: Response | undefined;
    void pending.then((value) => (response = value));
    try {
      await flush();
      expect(response).toBeDefined();
      expect(transcript[0]).toMatchObject({
        status: 200,
        request: '{"query":"hello"}',
        reply: '',
        bodyState: 'streaming',
      });
      expect(new Headers(fetchMock.mock.calls[0]?.[1]?.headers).get('X-Goog-User-Project')).toBe(
        'quota-project',
      );
      const reader = response!.body!.getReader();
      expect((await reader.read()).value).toEqual(chunk);
      reader.releaseLock();
    } finally {
      source.close();
      await pending;
    }
    await vi.waitFor(() => expect(transcript[0]?.reply).toBe('Hello'));
    expect(transcript[0]?.bodyState).toBe('complete');
  });

  it('preserves the HTTP receipt and records a failed body without hiding it as an empty call', async () => {
    let source!: ReadableStreamDefaultController<Uint8Array>;
    const body = new ReadableStream<Uint8Array>({ start: (controller) => (source = controller) });
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(body, { status: 200 })),
    );
    const transcript: TranscriptEntry[] = [];
    const pending = liveFetch(
      {},
      transcript,
    )(url).then(
      (response) => ({ response }),
      (error: unknown) => ({ error }),
    );
    await flush();
    source.error(new Error('stream disconnected'));
    const result = await pending;
    expect(result).toHaveProperty('response');
    if ('response' in result)
      await expect(result.response.text()).rejects.toThrow('stream disconnected');
    await vi.waitFor(() =>
      expect(transcript[0]).toMatchObject({ status: 200, bodyState: 'failed' }),
    );
  });
});
