import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';

/** A fake Word-like bridge: a selection that resolves to one text chunk, and a recording actuator. */
class FakeBridge implements DocBridge {
  readonly surface = 'word' as const;
  applied: ActuationRequest[] = [];

  getCapabilities(): CapabilityManifest {
    return {
      surface: 'word',
      contextKinds: ['selection'],
      actuations: [
        {
          kind: 'tracked-change',
          surface: 'word',
          title: 'Insert tracked change',
          reversible: true,
        },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([
      { id: 'word:selection', kind: 'selection', surface: 'word', title: 'Selection', live: true },
    ]);
  }
  resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    return Promise.resolve([
      { ref, value: { as: 'text', text: 'The SLA is 99.5%.', mimeType: 'text/markdown' } },
    ]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'para:3',
    });
  }
}

function streamOf(pieces: string[]): ReadableStream<Uint8Array> {
  const enc = new TextEncoder();
  let i = 0;
  return new ReadableStream({
    pull(c) {
      if (i < pieces.length) c.enqueue(enc.encode(pieces[i++]!));
      else c.close();
    },
  });
}

function geminiFetch() {
  const chunks = [
    {
      sessionInfo: { session: 'sess_1' },
      answer: { replies: [{ groundedContent: { content: { text: 'Below ' } } }] },
    },
    {
      answer: {
        state: 'SUCCEEDED',
        replies: [
          {
            groundedContent: {
              content: { text: 'the FSI floor.' },
              textGroundingMetadata: {
                references: [{ documentMetadata: { title: 'Vendor Policy', uri: 'https://x' } }],
              },
            },
          },
        ],
      },
    },
  ];
  return vi.fn(async () => new Response(streamOf([JSON.stringify(chunks)]), { status: 200 }));
}

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' };
const tokens = { getAccessToken: () => Promise.resolve('t') };
const unit = { connectors: [], surfaceContext: { kind: 'word' as const } };

async function collect(gen: AsyncGenerator<SseEvent>): Promise<SseEvent[]> {
  const out: SseEvent[] = [];
  for await (const e of gen) out.push(e);
  return out;
}

describe('AssistSession — the reusable loop', () => {
  it('auto-attaches bridge context, streams a grounded answer, captures session + citations', async () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, { unit, autoAttach: ['selection'] });

    const events = await collect(session.ask('Is the SLA below our floor?'));

    expect(session.context.size).toBe(1); // selection attached
    expect(events.map((e) => e.type)).toContain('token');
    expect(session.sessionId).toBe('sess_1');
    expect(session.sources[0]?.title).toBe('Vendor Policy');
  });

  it('applies an actuation through the bridge, stamping last-turn provenance', async () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, { unit, autoAttach: ['selection'] });
    await collect(session.ask('rewrite the SLA clause'));

    const result = await session.apply(
      'tracked-change',
      { text: 'The SLA is 99.9%.', target: { matchText: 'The SLA is 99.5%.' } },
      'change-1',
    );

    expect(result.ok).toBe(true);
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]).toMatchObject({
      kind: 'tracked-change',
      surface: 'word',
      changeId: 'change-1',
    });
    expect(bridge.applied[0]!.provenance?.identity).toBe('v.k@acme');
    expect(bridge.applied[0]!.provenance?.sources[0]?.title).toBe('Vendor Policy');
  });

  it('lets the caller attach/detach context explicitly', async () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, { unit });
    await session.attachContext(['selection']);
    expect(session.context.size).toBe(1);
    session.detach('word:selection');
    expect(session.context.size).toBe(0);
  });

  it('resumes a prior session id (cross-surface / reopen)', () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, { unit, resumeSessionId: 'sess_prior' });
    expect(session.sessionId).toBe('sess_prior');
  });
});

/** A fetch that records every streamAssist request body, returning a minimal valid stream. */
function recordingFetch(): { fetch: typeof fetch; bodies: Array<Record<string, unknown>> } {
  const bodies: Array<Record<string, unknown>> = [];
  const chunk = {
    sessionInfo: { session: 'sess_1' },
    answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text: 'ok' } } }] },
  };
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    bodies.push(JSON.parse(init?.body ?? '{}') as Record<string, unknown>);
    return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
  });
  return { fetch: fetchImpl as unknown as typeof fetch, bodies };
}

/** Pull the text parts out of a recorded streamAssist body. */
function partTexts(body: Record<string, unknown>): string[] {
  const query = body.query as { parts?: Array<{ text?: string }> } | undefined;
  return (query?.parts ?? []).map((p) => p.text ?? '').filter(Boolean);
}

describe('AssistSession — event-fed context model (construct → commit)', () => {
  it('folds a constructed brief into the next ask, then marks it resident (not re-sent)', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    // An event constructs context (a comment landed) — no model call yet.
    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'reconcile with SOW',
    });
    expect(bodies).toHaveLength(0); // fold is lazy — nothing sent
    expect(session.model.hasPending).toBe(true);

    // The first real question carries the brief.
    await collect(session.ask('what changed?'));
    expect(bodies).toHaveLength(1);
    expect(partTexts(bodies[0]!).some((t) => t.includes('reconcile with SOW'))).toBe(true);
    expect(session.model.hasPending).toBe(false); // now resident in the session

    // The next question does NOT re-send the already-committed brief.
    await collect(session.ask('and now?'));
    expect(bodies).toHaveLength(2);
    expect(partTexts(bodies[1]!).some((t) => t.includes('reconcile with SOW'))).toBe(false);
  });

  it('primes immediately at a checkpoint (meeting-ended), sending a context-only turn', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await session.ingest({ type: 'meeting-ended', id: 'mtg-9' });

    expect(bodies).toHaveLength(1); // primed now, before any user question
    expect(partTexts(bodies[0]!).some((t) => t.includes('mtg-9'))).toBe(true);
    expect(session.model.hasPending).toBe(false);
    expect(session.sessionId).toBe('sess_1'); // captured from the prime turn
  });

  it('ignores remote events — never narrates coauthor edits back into context', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'remote',
      commentId: 'k9',
      text: 'a coauthor note',
    });
    expect(session.model.hasPending).toBe(false);
    expect(bodies).toHaveLength(0);
  });
});
