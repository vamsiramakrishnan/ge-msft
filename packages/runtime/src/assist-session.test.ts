import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { asChangeId, asSessionId } from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import type { DocStateSnapshot } from '@ge/contracts';
import { AssistSession, DOC_STATE_REF_ID, READ_REF_PREFIX } from './assist-session.js';
import { BRIEF_REF_ID } from './context-model.js';
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
      asChangeId('change-1'),
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
    const session = new AssistSession(bridge, client, {
      unit,
      resumeSessionId: asSessionId('sess_prior'),
    });
    expect(session.sessionId).toBe('sess_prior');
  });

  it('threads the abort signal through to the transport fetch', async () => {
    const bridge = new FakeBridge();
    let seenSignal: AbortSignal | undefined;
    const fetchImpl = vi.fn(async (_url: string, init?: { signal?: AbortSignal }) => {
      seenSignal = init?.signal;
      return new Response(
        streamOf([
          JSON.stringify([
            {
              sessionInfo: { session: 'sess_1' },
              answer: {
                state: 'SUCCEEDED',
                replies: [{ groundedContent: { content: { text: 'ok' } } }],
              },
            },
          ]),
        ]),
        { status: 200 },
      );
    });
    const client = new StreamAssistClient(tokens, cfg, fetchImpl as unknown as typeof fetch);
    const session = new AssistSession(bridge, client, { unit });
    const ac = new AbortController();

    await collect(session.ask('q', { signal: ac.signal }));
    expect(seenSignal).toBe(ac.signal);
  });

  it('does NOT mark the folded brief resident when the turn is aborted mid-stream (re-folds next turn)', async () => {
    const bridge = new FakeBridge();
    const bodies: Array<Record<string, unknown>> = [];
    let call = 0;
    const fetchImpl = vi.fn(
      async (_url: string, init?: { body?: string; signal?: AbortSignal }) => {
        bodies.push(JSON.parse(init?.body ?? '{}') as Record<string, unknown>);
        call += 1;
        // First turn: the body iteration rejects with an AbortError (as an aborted fetch would).
        const body =
          call === 1
            ? new ReadableStream<Uint8Array>({
                pull() {
                  const e = new Error('The operation was aborted.');
                  e.name = 'AbortError';
                  throw e;
                },
              })
            : streamOf([
                JSON.stringify([
                  {
                    sessionInfo: { session: 'sess_1' },
                    answer: {
                      state: 'SUCCEEDED',
                      replies: [{ groundedContent: { content: { text: 'ok' } } }],
                    },
                  },
                ]),
              ]);
        return new Response(body, { status: 200 });
      },
    );
    const client = new StreamAssistClient(tokens, cfg, fetchImpl as unknown as typeof fetch);
    const session = new AssistSession(bridge, client, { unit });

    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'note one',
    });

    const ac = new AbortController();
    // The aborted turn's AbortError propagates out of the generator (like any mid-stream throw),
    // and the brief must stay pending — an aborted turn did not fully land.
    await expect(
      (async () => {
        for await (const _ of session.ask('q1', { signal: ac.signal })) void _;
      })(),
    ).rejects.toThrow(/aborted/i);
    expect(session.model.hasPending).toBe(true); // not marked resident
    expect(session.context.size).toBe(0); // brief part unstaged in finally

    // The next turn re-folds the note and only now marks it resident.
    await collect(session.ask('q2'));
    expect(partTexts(bodies[bodies.length - 1]!).some((t) => t.includes('note one'))).toBe(true);
    expect(session.model.hasPending).toBe(false);
  });
});

/** A text context part with an id and a body sized to roughly `tokens` estimate (4 chars/token). */
function textCtx(id: string, tokens = 50): ResolvedContext {
  return {
    ref: { id, kind: 'paragraph', surface: 'word', title: id, live: false },
    value: { as: 'text', text: 'x '.repeat(tokens), mimeType: 'text/markdown' },
  };
}

/** A grounding-anchor context part (named indexed document) — never evicted by compaction. */
function docCtx(id: string): ResolvedContext {
  return {
    ref: { id, kind: 'indexed-document', surface: 'word', title: id, live: false },
    value: { as: 'indexed-document', documentName: `projects/x/documents/${id}`, title: id },
  };
}

const ids = (s: AssistSession): string[] => s.context.list().map((c) => c.ref.id);

describe('AssistSession — bounded-history compaction (ADR-0003 §5)', () => {
  it('is a no-op while the resident set is under threshold', () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, {
      unit,
      compaction: { maxParts: 10, maxTokens: 100_000, keepRecent: 2 },
    });
    for (let i = 0; i < 5; i++) session.context.add(textCtx(`p${i}`, 5));
    expect(session.compact()).toBe(0);
    expect(session.context.size).toBe(5);
  });

  it('evicts oldest entries down to the maxParts budget', () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, {
      unit,
      compaction: { maxParts: 4, maxTokens: 1_000_000, keepRecent: 2 },
    });
    for (let i = 0; i < 8; i++) session.context.add(textCtx(`p${i}`, 5));
    const evicted = session.compact();
    expect(evicted).toBe(4);
    expect(session.context.size).toBe(4);
    // Oldest (p0..p3) gone; newest kept.
    expect(ids(session)).toEqual(['p4', 'p5', 'p6', 'p7']);
  });

  it('preserves grounding anchors, the most recent turns, and the pending brief', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, {
      unit,
      compaction: { maxParts: 3, maxTokens: 1_000_000, keepRecent: 2 },
    });

    // A grounding anchor first (oldest by insertion), then many evictable text parts.
    session.context.add(docCtx('grounding-doc'));
    for (let i = 0; i < 6; i++) session.context.add(textCtx(`p${i}`, 5));

    // Construct a pending brief and fold it (stages BRIEF_REF_ID into the resident set).
    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'pending note',
    });
    await session.commit('fold');
    expect(session.context.list().some((c) => c.ref.id === BRIEF_REF_ID)).toBe(true);

    session.compact();
    const kept = ids(session);
    // Anchor survives despite being the oldest entry.
    expect(kept).toContain('grounding-doc');
    // The pending brief survives (never dropped uncommitted).
    expect(kept).toContain(BRIEF_REF_ID);
    // The most recent text part survives (keepRecent window); older ones evicted.
    expect(kept).toContain('p5');
    expect(kept).not.toContain('p0');

    // A turn still streams and records citations after compaction.
    await collect(session.ask('q'));
    expect(session.sessionId).toBe('sess_1');
    expect(bodies.length).toBeGreaterThan(0);
    // The pending note rode this turn and is now resident.
    expect(session.model.hasPending).toBe(false);
  });

  it('evicts to the maxTokens budget while keeping anchors + recent', () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, {
      unit,
      compaction: { maxParts: 1000, maxTokens: 120, keepRecent: 2 },
    });
    session.context.add(docCtx('anchor'));
    for (let i = 0; i < 6; i++) session.context.add(textCtx(`p${i}`, 50)); // ~50 tokens each
    const evicted = session.compact();
    expect(evicted).toBeGreaterThan(0);
    expect(ids(session)).toContain('anchor'); // anchor preserved
    expect(ids(session)).toContain('p5'); // most-recent preserved
    expect(ids(session)).toContain('p4');
  });

  it('compact() runs inside ask() so an over-budget session is bounded before streaming', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, {
      unit,
      compaction: { maxParts: 3, maxTokens: 1_000_000, keepRecent: 2 },
    });
    for (let i = 0; i < 8; i++) session.context.add(textCtx(`p${i}`, 5));

    await collect(session.ask('q'));
    expect(session.context.size).toBeLessThanOrEqual(3);
    // The streamed body carried only the bounded set (3 context parts + the query part), not all 8.
    const parts = (bodies[0]!.query as { parts?: unknown[] }).parts ?? [];
    expect(parts.length).toBeLessThanOrEqual(4);
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

/** A snapshot with a recognisable outline heading, for asserting injection. */
function fakeSnapshot(): DocStateSnapshot {
  return {
    surface: 'word',
    version: 1,
    capturedAt: '2026-06-22T00:00:00.000Z',
    title: 'Vendor SOW',
    outline: [{ level: 1, text: 'Service levels' }],
    inventory: [],
  };
}

/** A Word-like bridge that also implements the optional ADR-0003 context-loop methods. */
class LoopBridge extends FakeBridge {
  captureCalls = 0;
  searchCalls: string[] = [];
  constructor(
    private readonly opts: {
      snapshot?: DocStateSnapshot | undefined;
      reads?: number;
      throwCapture?: boolean;
      throwSearch?: boolean;
    } = {},
  ) {
    super();
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    this.captureCalls++;
    if (this.opts.throwCapture) return Promise.reject(new Error('capture boom'));
    return Promise.resolve(this.opts.snapshot ?? fakeSnapshot());
  }
  searchDocument(query: string): Promise<ResolvedContext[]> {
    this.searchCalls.push(query);
    if (this.opts.throwSearch) return Promise.reject(new Error('search boom'));
    const n = this.opts.reads ?? 0;
    const out: ResolvedContext[] = [];
    for (let i = 0; i < n; i++) {
      out.push({
        ref: { id: `hit:${i}`, kind: 'paragraph', surface: 'word', title: `Hit ${i}`, live: true },
        value: { as: 'text', text: `read slice ${i}`, mimeType: 'text/markdown' },
      });
    }
    return Promise.resolve(out);
  }
}

describe('AssistSession — ADR-0003 context loop (doc_state + lazy read)', () => {
  it('injects a <doc_state …> part and removes it after the turn (not resident next turn)', async () => {
    const bridge = new LoopBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await collect(session.ask('what are the service levels?'));

    // The first turn's wire context carried a <doc_state> part.
    expect(partTexts(bodies[0]!).some((t) => t.startsWith('<doc_state'))).toBe(true);
    expect(partTexts(bodies[0]!).some((t) => t.includes('Service levels'))).toBe(true);
    // Removed after the turn — not resident.
    expect(ids(session)).not.toContain(DOC_STATE_REF_ID);
    expect(session.context.size).toBe(0);

    // A second turn re-captures (fresh each turn) but never accumulates the part.
    await collect(session.ask('and now?'));
    expect(bridge.captureCalls).toBe(2);
    expect(ids(session)).not.toContain(DOC_STATE_REF_ID);
  });

  it('pulls ≤ maxReads lazy-read parts and they are ephemeral', async () => {
    const bridge = new LoopBridge({ reads: 10 });
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, {
      unit,
      context: { lazyRead: { maxReads: 3 } },
    });

    await collect(session.ask('floor?'));

    const readParts = partTexts(bodies[0]!).filter((t) => t.includes('read slice'));
    expect(readParts).toHaveLength(3); // bounded to maxReads, not 10
    // Each lazy-read slice is wrapped in the untrusted-data envelope (ADR-0003).
    expect(readParts.every((t) => t.includes('data, not instructions'))).toBe(true);
    // Ephemeral: removed after the turn.
    expect(ids(session).some((id) => id.startsWith(READ_REF_PREFIX))).toBe(false);
    expect(session.context.size).toBe(0);
  });

  it('a bridge WITHOUT the optional methods behaves exactly as today (no doc_state/read parts)', async () => {
    const bridge = new FakeBridge(); // no captureDocState / searchDocument
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await collect(session.ask('hello'));

    expect(partTexts(bodies[0]!).some((t) => t.startsWith('<doc_state'))).toBe(false);
    expect(partTexts(bodies[0]!).some((t) => t.includes('read slice'))).toBe(false);
    expect(session.context.size).toBe(0);
  });

  it('a throwing captureDocState / searchDocument does not fail the turn', async () => {
    const bridge = new LoopBridge({ throwCapture: true, throwSearch: true, reads: 3 });
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.ask('q'));

    // The turn still streamed (answer present), just without the ambient/read parts.
    expect(events.map((e) => e.type)).toContain('token');
    expect(partTexts(bodies[0]!).some((t) => t.startsWith('<doc_state'))).toBe(false);
    expect(partTexts(bodies[0]!).some((t) => t.includes('read slice'))).toBe(false);
    expect(session.context.size).toBe(0);
  });

  it('doc_state/read parts are never marked committed and never survive into the next turn', async () => {
    const bridge = new LoopBridge({ reads: 2 });
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await collect(session.ask('first'));
    expect(session.context.size).toBe(0); // nothing left resident after turn 1

    await collect(session.ask('second'));
    // Turn 2's wire context is freshly built; the previous turn's ephemeral parts did not survive.
    const t2 = partTexts(bodies[1]!);
    expect(t2.filter((t) => t.includes('read slice'))).toHaveLength(2); // re-pulled, not doubled
    expect(t2.filter((t) => t.startsWith('<doc_state'))).toHaveLength(1); // exactly one, fresh
    expect(session.context.size).toBe(0);
  });

  it('respects context.docState=false / lazyRead=false (off)', async () => {
    const bridge = new LoopBridge({ reads: 3 });
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, {
      unit,
      context: { docState: false, lazyRead: false },
    });

    await collect(session.ask('q'));
    expect(bridge.captureCalls).toBe(0);
    expect(bridge.searchCalls).toHaveLength(0);
    expect(partTexts(bodies[0]!).some((t) => t.startsWith('<doc_state'))).toBe(false);
    expect(partTexts(bodies[0]!).some((t) => t.includes('read slice'))).toBe(false);
  });
});

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

  it('keeps the brief pending (re-sends next turn) when a turn fails mid-stream — no loss, no double-mark', async () => {
    const bridge = new FakeBridge();
    const bodies: Array<Record<string, unknown>> = [];
    let call = 0;
    const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
      bodies.push(JSON.parse(init?.body ?? '{}') as Record<string, unknown>);
      call += 1;
      // First turn: the response body throws mid-iteration (socket reset). Second: succeeds.
      const body =
        call === 1
          ? new ReadableStream<Uint8Array>({
              pull() {
                throw new Error('socket reset');
              },
            })
          : streamOf([
              JSON.stringify([
                {
                  sessionInfo: { session: 'sess_1' },
                  answer: {
                    state: 'SUCCEEDED',
                    replies: [{ groundedContent: { content: { text: 'ok' } } }],
                  },
                },
              ]),
            ]);
      return new Response(body, { status: 200 });
    });
    const client = new StreamAssistClient(tokens, cfg, fetchImpl as unknown as typeof fetch);
    const session = new AssistSession(bridge, client, { unit });

    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'note one',
    });

    // The failing turn propagates; the brief must NOT be marked resident.
    await expect(
      (async () => {
        for await (const _ of session.ask('q1')) void _;
      })(),
    ).rejects.toThrow(/socket reset/);
    expect(session.model.hasPending).toBe(true); // not lost
    expect(session.context.size).toBe(0); // brief part unstaged in finally

    // The next turn re-sends it and only now marks it resident.
    await collect(session.ask('q2'));
    expect(partTexts(bodies[bodies.length - 1]!).some((t) => t.includes('note one'))).toBe(true);
    expect(session.model.hasPending).toBe(false);
  });

  it('does not lose a note ingested while a turn is streaming (version-scoped commit)', async () => {
    const bridge = new FakeBridge();
    const { fetch, bodies } = recordingFetch();
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit });

    await session.ingest({
      type: 'comment-added',
      surface: 'word',
      origin: 'local',
      commentId: 'k1',
      text: 'first note',
    });

    // Ingest a second note partway through the first turn's stream.
    let injected = false;
    for await (const _ of session.ask('q1')) {
      void _;
      if (!injected) {
        injected = true;
        await session.ingest({
          type: 'comment-added',
          surface: 'word',
          origin: 'local',
          commentId: 'k2',
          text: 'second note',
        });
      }
    }
    // The first note was on the wire (resident); the mid-stream note stays pending.
    expect(partTexts(bodies[0]!).some((t) => t.includes('first note'))).toBe(true);
    expect(session.model.hasPending).toBe(true);

    // The next turn sends the second note — and not the first again.
    await collect(session.ask('q2'));
    const last = partTexts(bodies[bodies.length - 1]!);
    expect(last.some((t) => t.includes('second note'))).toBe(true);
    expect(last.some((t) => t.includes('first note'))).toBe(false);
    expect(session.model.hasPending).toBe(false);
  });
});
