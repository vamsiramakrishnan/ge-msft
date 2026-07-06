import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ProvenancePayload,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { asChangeId, asSessionId } from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import type { DocStateSnapshot } from '@ge/contracts';
import { AssistSession, DOC_STATE_REF_ID, READ_REF_PREFIX } from './assist-session.js';
import type { CommandLoopEvent, PlanEffect } from './assist-session.js';
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

  it('applies an actuation through the bridge, stamping the EXPLICIT provenance the caller passes', async () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(tokens, cfg, geminiFetch() as never);
    const session = new AssistSession(bridge, client, { unit, autoAttach: ['selection'] });

    // Finding #4: provenance is turn-scoped. The caller captures the turn's `provenance` event and
    // passes it back into `apply` EXPLICITLY — `apply` never reads an ambient instance field.
    let captured: ProvenancePayload | undefined;
    for await (const ev of session.ask('rewrite the SLA clause')) {
      if (ev.type === 'provenance') captured = ev.payload;
    }
    expect(captured?.sources[0]?.title).toBe('Vendor Policy');

    const result = await session.apply(
      'tracked-change',
      { text: 'The SLA is 99.9%.', target: { matchText: 'The SLA is 99.5%.' } },
      asChangeId('change-1'),
      captured,
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

  it('Finding #4: provenance is turn-scoped — a later provenance-less turn clears the prior turn’s, so a write inherits NONE of turn A’s', async () => {
    const bridge = new FakeBridge();
    // Turn A emits provenance; turn B is policy-BLOCKED, so it streams NO provenance event at all.
    const fetchImpl = vi.fn(async () => {
      const callIndex = fetchImpl.mock.calls.length; // 1-based after this call records
      if (callIndex === 1) {
        const chunk = {
          sessionInfo: { session: 'sess_A' },
          answer: {
            state: 'SUCCEEDED',
            replies: [
              {
                groundedContent: {
                  content: { text: 'A' },
                  textGroundingMetadata: {
                    references: [
                      { documentMetadata: { title: 'Vendor Policy', uri: 'https://x' } },
                    ],
                  },
                },
              },
            ],
          },
        };
        return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
      }
      // Turn B: a Model-Armor BLOCK verdict suppresses the answer AND the provenance for the turn.
      const blocked = {
        answer: { customerPolicyEnforcementResult: { verdict: 'BLOCK' } },
      };
      return new Response(streamOf([JSON.stringify([blocked])]), { status: 200 });
    });
    const client = new StreamAssistClient(tokens, cfg, fetchImpl as unknown as typeof fetch);
    const session = new AssistSession(bridge, client, { unit, autoAttach: ['selection'] });

    // Turn A: a grounded ask that DOES emit provenance — captured turn-scoped.
    const aEvents = await collect(session.ask('rewrite the SLA clause'));
    expect(aEvents.some((e) => e.type === 'provenance')).toBe(true);

    // Turn B: a provenance-LESS turn. It must CLEAR turn A's provenance from the turn-scoped slot.
    const bEvents = await collect(session.ask('a follow-up'));
    expect(bEvents.some((e) => e.type === 'provenance')).toBe(false);

    // A write made now (no explicit provenance) must inherit NOTHING — not turn A's leftover.
    const result = await session.apply(
      'tracked-change',
      { text: 'x', target: { matchText: 'The SLA is 99.5%.' } },
      asChangeId('change-2'),
    );

    expect(result.ok).toBe(true);
    expect(bridge.applied[0]!.provenance).toBeUndefined();
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

/* ───────────────────────── ADR-0005 composition in the loop ───────────────── */

/** An Excel-like bridge whose `readRange` returns a GFM table (the shape Excel reads produce). */
class ComposeBridge implements DocBridge {
  readonly surface = 'excel' as const;
  readRangeCalls: string[] = [];
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'excel',
      contextKinds: ['range'],
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
        { kind: 'insert-chart', surface: 'excel', title: 'Insert chart', reversible: true },
        // Advertised so a composition-parity test can resolve `slide`'s bulletsExpr through the loop.
        { kind: 'insert-slide', surface: 'excel', title: 'Insert slide', reversible: true },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
  readRange(a1: string): Promise<ResolvedContext[]> {
    this.readRangeCalls.push(a1);
    const gfm = '| region | amount |\n| --- | --- |\n| East | 100 |\n| West | 250 |\n| East | 50 |';
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: { as: 'text', text: gfm, mimeType: 'text/markdown' },
      },
    ]);
  }
}

/**
 * A streamAssist fetch scripted per call: each element is the text the model "emits" that turn.
 * The runtime feeds the previous turn's ```result``` block back, so a multi-turn composition test
 * can assert the evaluated Value arrives in the next turn's query.
 */
function scriptedFetch(turns: string[]): {
  fetch: typeof fetch;
  bodies: Array<Record<string, unknown>>;
} {
  const bodies: Array<Record<string, unknown>> = [];
  let call = 0;
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    bodies.push(JSON.parse(init?.body ?? '{}') as Record<string, unknown>);
    const text = turns[Math.min(call, turns.length - 1)] ?? '```cmd\ndone\n```';
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
  });
  return { fetch: fetchImpl as unknown as typeof fetch, bodies };
}

async function collectLoop(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const out: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) out.push(e);
  return out;
}

describe('AssistSession.runCommands — ADR-0005 composition (pure)', () => {
  it('mounts the configured command skill route for command-loop turns', async () => {
    const bridge = new ComposeBridge();
    const { fetch, bodies } = scriptedFetch(['```cmd\ndone\n```']);
    const client = new StreamAssistClient(
      tokens,
      {
        ...cfg,
        commandSkills: ['projects/proj/locations/global/agents/7404511736383961129'],
        commandSkillMentions: [{ label: 'm365-surface-commander', uri: '7404511736383961129' }],
        plannerSkills: ['projects/proj/locations/global/agents/17573173582293271726'],
        plannerSkillMentions: [{ label: 'm365-command-planner', uri: '17573173582293271726' }],
      },
      fetch,
    );
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    await collectLoop(session.runCommands('/visualize @this'));

    const body = bodies[0]!;
    const query = (body.query as { text?: string }).text ?? '';
    expect(body.skillsSpec).toEqual({
      skills: [{ name: 'projects/proj/locations/global/agents/7404511736383961129' }],
    });
    expect(query).toContain('[m365-surface-commander](mention://?uri=7404511736383961129)');
    expect(query).not.toContain('m365-command-planner');
  });

  it('evaluates a read|filter|sum pipeline and feeds the Value back in the next turn', async () => {
    const bridge = new ComposeBridge();
    const { fetch, bodies } = scriptedFetch([
      '```cmd\nread Sales!A1:B5 | filter region=East | sum amount\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('total East'));

    // The pipeline evaluated to East amounts 100 + 50 = 150.
    const exprEvent = events.find((e) => e.type === 'expr-result') as
      | Extract<CommandLoopEvent, { type: 'expr-result' }>
      | undefined;
    expect(exprEvent?.result).toEqual({ kind: 'number', value: 150 });

    // Turn 2's query carried the rendered value back in the ```result``` block.
    const turn2 = (bodies[1]!.query as { text?: string }).text ?? '';
    expect(turn2).toContain('150');
  });

  it('persists a $var binding across turns within the loop', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read Sales!A1:B5\n```',
      '```cmd\n$t | count\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('count rows'));

    const exprEvents = events.filter((e) => e.type === 'expr-result') as Array<
      Extract<CommandLoopEvent, { type: 'expr-result' }>
    >;
    // The second expression resolved the $var bound in the first turn → 3 rows.
    expect(exprEvents[1]?.result).toEqual({ kind: 'number', value: 3 });
  });

  it('rejects a pipe-into-effect with the Phase-1 corrective (no write)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nread Sales!A1:B5 | set Sales!F2 =1\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('try to write'));
    const exprEvent = events.find((e) => e.type === 'expr-result') as
      | Extract<CommandLoopEvent, { type: 'expr-result' }>
      | undefined;
    expect(exprEvent?.result).toMatchObject({
      error: expect.stringContaining("can't be composed"),
    });
    // No write-result event — composition never actuated.
    expect(events.some((e) => e.type === 'write-result')).toBe(false);
  });
});

/* ───────────────── ADR-0005 Phase 2 — gated effect composition (plan) ───────── */

const planEvents = (
  events: Array<SseEvent | CommandLoopEvent>,
): Array<Extract<CommandLoopEvent, { type: 'plan-preview' }>> =>
  events.filter((e) => e.type === 'plan-preview') as Array<
    Extract<CommandLoopEvent, { type: 'plan-preview' }>
  >;

const writeResults = (
  events: Array<SseEvent | CommandLoopEvent>,
): Array<Extract<CommandLoopEvent, { type: 'write-result' }>> =>
  events.filter((e) => e.type === 'write-result') as Array<
    Extract<CommandLoopEvent, { type: 'write-result' }>
  >;

describe('AssistSession.runCommands — ADR-0005 Phase 2 (gated effect composition)', () => {
  it('resolves an effect-arg expression at dry-run to a concrete number write', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      // Bind a table, then write a composed value into a cell — the keystone connection.
      '```cmd\nlet $t = read Sales!A1:B5\n```',
      '```cmd\nset Summary!B2 = ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let previewed: PlanEffect[] | undefined;
    const events = await collectLoop(
      session.runCommands('write the total', {
        approvePlan: (effects) => {
          previewed = effects;
          return true;
        },
      }),
    );

    // The plan-preview carried the dry-run effect with the RESOLVED value: 100 + 250 + 50 = 400.
    const preview = planEvents(events);
    expect(preview).toHaveLength(1);
    expect(previewed?.[0]?.request).toMatchObject({
      kind: 'write-cells',
      params: { target: { range: 'Summary!B2' }, cells: [['400']] },
    });
    expect(previewed?.[0]?.command).toBe('set Summary!B2 = ($t | sum amount)');
    // The approval card sees the CONCRETE resolved value, not just the formula (security: the human
    // approves the value that will land, derived from doc content, not an opaque expression).
    expect(previewed?.[0]?.dryRun).toEqual({ target: 'Summary!B2', resolved: '400' });
    // Approved → it actuated exactly once.
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]?.params.cells).toEqual([['400']]);
    // A grounded turn emits provenance (streamAssist), so the stamped write IS attributed.
    expect(bridge.applied[0]?.provenance).toBeDefined();
  });

  it('resolves a slide bulletsExpr (table → bullets) at dry-run', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read Sales!A1:B5\n```',
      // The table's rows become bullets — composition parity for a surface verb.
      '```cmd\nslide "Regions" ($t | select region,amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let previewed: PlanEffect[] | undefined;
    await collectLoop(
      session.runCommands('build the slide', {
        approvePlan: (effects) => {
          previewed = effects;
          return true;
        },
      }),
    );

    // Each row of the projected table became one bullet (cells joined by " · ").
    expect(previewed?.[0]?.request).toMatchObject({
      kind: 'insert-slide',
      params: { slide: { title: 'Regions', bullets: ['East · 100', 'West · 250', 'East · 50'] } },
    });
    // The approval card previews the resolved bullets, not the opaque expression.
    expect(previewed?.[0]?.dryRun).toEqual({
      target: 'Regions',
      resolved: '• East · 100  • West · 250  • East · 50',
    });
    expect(bridge.applied[0]?.params.slide?.bullets).toEqual([
      'East · 100',
      'West · 250',
      'East · 50',
    ]);
  });

  it('resolves a spill valueExpr (table → cell grid) at dry-run (ADR-0007 §3)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read Sales!A1:B5\n```',
      // The projected table becomes a grid: a header row + one row per data row.
      '```cmd\nspill Report!A1 = ($t | select region,amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let previewed: PlanEffect[] | undefined;
    await collectLoop(
      session.runCommands('materialize the table', {
        approvePlan: (effects) => {
          previewed = effects;
          return true;
        },
      }),
    );

    // spill → write-cells with the header row first, then the data rows.
    expect(previewed?.[0]?.request).toMatchObject({
      kind: 'write-cells',
      params: {
        target: { range: 'Report!A1' },
        cells: [
          ['region', 'amount'],
          ['East', '100'],
          ['West', '250'],
          ['East', '50'],
        ],
      },
    });
    expect(bridge.applied[0]?.params.cells?.[0]).toEqual(['region', 'amount']);
  });

  it('runs a literal grid as one write-cells plan effect', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\ngrid Report!A1:B2 = "Region\\tRevenue\\nEast\\t100"\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let previewed: PlanEffect[] | undefined;
    await collectLoop(
      session.runCommands('materialize a literal grid', {
        approvePlan: (effects) => {
          previewed = effects;
          return true;
        },
      }),
    );

    expect(previewed).toHaveLength(1);
    expect(previewed?.[0]?.request).toMatchObject({
      kind: 'write-cells',
      params: {
        target: { range: 'Report!A1:B2' },
        cells: [
          ['Region', 'Revenue'],
          ['East', '100'],
        ],
      },
    });
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]?.params.cells).toEqual([
      ['Region', 'Revenue'],
      ['East', '100'],
    ]);
  });

  it('rejects a spill whose expression resolves to a scalar (composition guard, ADR-0007)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read Sales!A1:B5\n```',
      // A scalar terminal (`sum`) is NOT a table — spill must reject it (dual of set's table guard).
      '```cmd\nspill Report!A1 = ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('try a scalar spill', { approvePlan: () => true }),
    );
    expect(planEvents(events)).toHaveLength(0);
    expect(bridge.applied).toHaveLength(0); // nothing actuated
    const cmdErr = events.find((e) => e.type === 'command' && 'error' in e.compiled) as
      | Extract<CommandLoopEvent, { type: 'command' }>
      | undefined;
    expect(cmdErr?.compiled).toMatchObject({
      error: expect.stringContaining('spill needs a table'),
    });
  });

  it('type-check rejects an unsupported verb before execution (no write)', async () => {
    // Word-only verb `suggest` on an Excel surface that does not advertise tracked-change.
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch(['```cmd\nsuggest "a" => "b"\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('try unsupported', { approvePlan: () => true }),
    );
    // No plan-preview (nothing resolvable), no actuation; the model got a corrective in ```result```.
    expect(planEvents(events)).toHaveLength(0);
    expect(bridge.applied).toHaveLength(0);
    const cmdErr = events.find((e) => e.type === 'command' && 'error' in e.compiled) as
      | Extract<CommandLoopEvent, { type: 'command' }>
      | undefined;
    expect(cmdErr?.compiled).toMatchObject({ error: expect.stringContaining('not supported') });
  });

  it('type-check rejects an effect with an unbound $var before the gate (no write)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nset B2 = ($missing | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('write', { approvePlan: () => true }));
    // The unbound $var fails at dry-run resolution → no plan effect, no actuation.
    expect(planEvents(events)).toHaveLength(0);
    expect(bridge.applied).toHaveLength(0);
    const cmdErr = events.find((e) => e.type === 'command' && 'error' in e.compiled) as
      | Extract<CommandLoopEvent, { type: 'command' }>
      | undefined;
    expect(cmdErr?.compiled).toMatchObject({ error: expect.stringContaining('unbound') });
  });

  it('dry-run produces the effect-set without actuating; reject → nothing actuates', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\nset A2 2\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let sawEffects = 0;
    const events = await collectLoop(
      session.runCommands('two writes', {
        approvePlan: (effects) => {
          sawEffects = effects.length;
          return false; // reject the whole plan
        },
      }),
    );

    // The plan was previewed with BOTH effects (dry-run resolved them)…
    expect(sawEffects).toBe(2);
    expect(planEvents(events)[0]?.effects).toHaveLength(2);
    // …but the reject blocked every one — zero actuations, each a corrective plan_unapproved.
    expect(bridge.applied).toHaveLength(0);
    const writes = writeResults(events);
    expect(writes).toHaveLength(2);
    expect(writes.every((w) => w.result.error?.code === 'plan_unapproved')).toBe(true);
  });

  it('approve → every effect is gated + actuated (one approval for the whole set)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nset A1 1\nset A2 2\nset A3 3\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let calls = 0;
    await collectLoop(
      session.runCommands('three writes', {
        approvePlan: () => {
          calls += 1;
          return true;
        },
      }),
    );
    // ONE plan approval, three gated actuations.
    expect(calls).toBe(1);
    expect(bridge.applied).toHaveLength(3);
  });

  it('does not accept done batched after a write in the same command block', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nchart bar \'Project schedule\'!B5:D30 title="Task Progress"\ndone\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('chart then done', { approvePlan: () => true }),
    );

    const previews = planEvents(events);
    expect(previews).toHaveLength(1);
    expect(previews[0]?.effects[0]?.request).toMatchObject({
      kind: 'insert-chart',
      params: {
        chart: {
          chartType: 'bar',
          sourceRange: "'Project schedule'!B5:D30",
          title: 'Task Progress',
        },
      },
    });
    expect(bridge.applied).toHaveLength(1);
    expect(events.some((e) => e.type === 'done' && 'turn' in e && e.turn === 1)).toBe(false);
    expect(events.some((e) => e.type === 'done' && 'turn' in e && e.turn === 2)).toBe(true);
  });

  it('no approver ⇒ the whole plan is blocked (fail-closed); nothing actuates', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('write')); // no approvePlan / approveWrite
    expect(bridge.applied).toHaveLength(0);
    expect(writeResults(events)[0]?.result.error?.code).toBe('plan_unapproved');
  });

  it('back-compat: approveWrite (no approvePlan) falls back to ADR-0004 per-write approval', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\nset A2 2\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let perWrite = 0;
    await collectLoop(
      session.runCommands('two writes', {
        approveWrite: () => {
          perWrite += 1;
          return true;
        },
      }),
    );
    // Per-write approver called once per effect (ADR-0004 Track A), both actuated.
    expect(perWrite).toBe(2);
    expect(bridge.applied).toHaveLength(2);
  });

  it('an effect-arg that resolves to a table is rejected (scalar required, no write)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      // No scalar terminal → the pipeline value is a table; writing it into one cell is rejected.
      '```cmd\nset B2 = (read Sales!A1:B5 | filter region=East)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('write', { approvePlan: () => true }));
    expect(planEvents(events)).toHaveLength(0);
    expect(bridge.applied).toHaveLength(0);
    const cmdErr = events.find((e) => e.type === 'command' && 'error' in e.compiled) as
      | Extract<CommandLoopEvent, { type: 'command' }>
      | undefined;
    expect(cmdErr?.compiled).toMatchObject({ error: expect.stringContaining('scalar') });
  });

  it('the per-turn write cap holds (capped effects never reach the gate)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nset A1 1\nset A2 2\nset A3 3\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('three writes', { approvePlan: () => true, maxWritesPerTurn: 2 }),
    );
    expect(bridge.applied).toHaveLength(2); // third capped, never planned/actuated
    expect(planEvents(events)[0]?.effects).toHaveLength(2);
    expect(events.some((e) => e.type === 'capped')).toBe(true);
  });
});

/* ───────────────── ADR-0005 Phase 3 — named skills (def / call → plan) ───────── */

const skillExpanded = (
  events: Array<SseEvent | CommandLoopEvent>,
): Array<Extract<CommandLoopEvent, { type: 'skill-expanded' }>> =>
  events.filter((e) => e.type === 'skill-expanded') as Array<
    Extract<CommandLoopEvent, { type: 'skill-expanded' }>
  >;

describe('AssistSession.runCommands — ADR-0005 Phase 3 (named skills)', () => {
  it('a def registers (no execution) → a skill-registered event, nothing actuates', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\ndef writeTotal($a $b):\n  set $b = ($a | sum amount)\nend\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('define a skill', { approvePlan: () => true }),
    );

    const reg = events.find((e) => e.type === 'skill-registered') as
      | Extract<CommandLoopEvent, { type: 'skill-registered' }>
      | undefined;
    expect(reg?.result.ok).toBe(true);
    expect(reg?.name).toBe('writeTotal');
    // A def is not an effect — no plan, no actuation.
    expect(bridge.applied).toHaveLength(0);
    expect(events.some((e) => e.type === 'write-result')).toBe(false);
  });

  it('a call expands, binds args, and runs as ONE plan gated by approvePlan (approve → actuates)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      // Turn 1: define a 2-param skill whose body reads a range and writes a composed sum.
      '```cmd\ndef writeTotal($src $dst):\n  set $dst = (read $src | sum amount)\nend\n```',
      // Turn 2: call it — the expansion is the plan.
      '```cmd\nwriteTotal Sales!A1:B5 Summary!B2\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let previewed: PlanEffect[] | undefined;
    const events = await collectLoop(
      session.runCommands('run the skill', {
        approvePlan: (effects) => {
          previewed = effects;
          return true;
        },
      }),
    );

    // The call expanded with $src/$dst substituted.
    const exp = skillExpanded(events)[0];
    expect(exp?.name).toBe('writeTotal');
    expect(exp?.lines).toEqual(['set Summary!B2 = (read Sales!A1:B5 | sum amount)']);

    // The expansion formed ONE plan: a single preview, single approval, the composed sum resolved.
    expect(planEvents(events)).toHaveLength(1);
    expect(previewed).toHaveLength(1);
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]).toMatchObject({
      kind: 'write-cells',
      params: { target: { range: 'Summary!B2' }, cells: [['400']] }, // 100+250+50
    });
  });

  it('rejecting the plan blocks the WHOLE expansion — nothing actuates', async () => {
    const bridge = new ComposeBridge();
    // The body binds a read, then writes the composed count into two cells — two effects, one plan.
    const { fetch } = scriptedFetch([
      '```cmd\ndef two($src $c1 $c2):\n  let $t = read $src\n  set $c1 = ($t | count)\n  set $c2 = ($t | count)\nend\n```',
      '```cmd\ntwo Sales!A1:B5 Summary!B2 Summary!B3\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('run + reject', { approvePlan: () => false }),
    );

    // Two effects previewed in one plan, but the reject blocked every one — zero actuations.
    expect(planEvents(events)[0]?.effects).toHaveLength(2);
    expect(bridge.applied).toHaveLength(0);
    const blocked = writeResults(events);
    expect(blocked).toHaveLength(2);
    expect(blocked.every((w) => w.result.error?.code === 'plan_unapproved')).toBe(true);
  });

  it('an undefined-name call is corrective (not actuated)', async () => {
    const bridge = new ComposeBridge();
    // `mystery` is not a built-in verb and not registered → unknown-verb command error.
    const { fetch } = scriptedFetch(['```cmd\nmystery A1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(
      session.runCommands('call missing', { approvePlan: () => true }),
    );
    expect(bridge.applied).toHaveLength(0);
    expect(events.some((e) => e.type === 'write-result')).toBe(false);
  });

  it('an arity-mismatched call is corrective — nothing actuates', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\ndef f($a $b):\n  set $a = ($a | count)\nend\n```',
      '```cmd\nf Sales!A1:B5\n```', // one arg, expects two
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('bad arity', { approvePlan: () => true }));
    const exp = skillExpanded(events)[0];
    expect(exp?.lines).toEqual([]); // expansion failed → no lines
    expect(bridge.applied).toHaveLength(0);
  });

  it('a def whose name shadows a built-in is rejected at register time', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\ndef set($a):\n  read $a\nend\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('shadow', { approvePlan: () => true }));
    // The parser rejects `def set(...)` as a corrective error entry (not a skill-registered ok).
    const reg = events.find((e) => e.type === 'skill-registered') as
      | Extract<CommandLoopEvent, { type: 'skill-registered' }>
      | undefined;
    // Either it never registered, or it registered with ok:false — never a usable `set` override.
    expect(reg?.result.ok ?? false).toBe(false);
  });

  it('SECURITY: a self-recursive skill is bounded (no stack overflow, never actuates)', async () => {
    const bridge = new ComposeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\ndef loop($a):\n  loop $a\nend\n```',
      '```cmd\nloop Sales!A1\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    // Terminates (depth guard) and surfaces a corrective; zero actuations.
    const events = await collectLoop(session.runCommands('recurse', { approvePlan: () => true }));
    expect(bridge.applied).toHaveLength(0);
    expect(
      events.some(
        (e) => e.type === 'capped' || (e.type === 'skill-expanded' && e.lines.length === 0),
      ),
    ).toBe(true);
  });

  it('SECURITY: a skill expansion cannot exceed the per-turn command cap', async () => {
    const bridge = new ComposeBridge();
    // A body of 4 reads; cap the turn to 2 commands — the call expands but is budget-bounded.
    const { fetch } = scriptedFetch([
      '```cmd\ndef many($a):\n  read $a\n  read $a\n  read $a\n  read $a\nend\n```',
      '```cmd\nmany Sales!A1:B5\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    await collectLoop(
      session.runCommands('budget', { approvePlan: () => true, maxCommandsPerTurn: 2 }),
    );
    // The call itself costs 1 budget unit; only 1 expanded read runs before the budget is exhausted.
    // (readRange is the read port — at most `maxCommandsPerTurn` reads can fire across the expansion.)
    expect(bridge.readRangeCalls.length).toBeLessThanOrEqual(2);
    expect(bridge.readRangeCalls.length).toBeGreaterThan(0);
  });
});

/** A fetch that returns one streamAssist chunk carrying `text` (for the planner pre-stage). */
function fetchText(text: string) {
  const chunks = [
    {
      sessionInfo: { session: 'sess_p' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    },
  ];
  return vi.fn(async () => new Response(streamOf([JSON.stringify(chunks)]), { status: 200 }));
}

describe('AssistSession.plan — the planner pre-stage (§F)', () => {
  it('streams one ```plan turn and returns the parsed CommandPlan', async () => {
    const bridge = new FakeBridge();
    const planText =
      '```plan\nintent rewrite\nsurface word\nstep rewrite the SLA to 99.9% as a tracked change\nexclude the indemnity clause\n```';
    const client = new StreamAssistClient(tokens, cfg, fetchText(planText) as never);
    const session = new AssistSession(bridge, client, { unit });

    const { plan, errors, needsClarification } = await session.plan(
      'rewrite the SLA but leave indemnity',
    );

    expect(errors).toEqual([]);
    expect(needsClarification).toBe(false);
    expect(plan?.intent).toBe('rewrite');
    expect(plan?.surface).toBe('word');
    expect(plan?.steps).toEqual(['rewrite the SLA to 99.9% as a tracked change']);
    expect(plan?.excludes).toEqual(['the indemnity clause']);
  });

  it('surfaces a clarify-only plan as needsClarification (no steps required)', async () => {
    const bridge = new FakeBridge();
    const client = new StreamAssistClient(
      tokens,
      cfg,
      fetchText(
        '```plan\nintent rewrite\nsurface word\nclarify which section — §4 or §5?\n```',
      ) as never,
    );
    const session = new AssistSession(bridge, client, { unit });
    const { plan, needsClarification } = await session.plan('rewrite the section');
    expect(needsClarification).toBe(true);
    expect(plan?.clarify).toEqual(['which section — §4 or §5?']);
  });
});
