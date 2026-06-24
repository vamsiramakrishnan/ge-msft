// @vitest-environment node
import { describe, it, expect } from 'vitest';
import {
  asChangeId,
  type ActuationRequest,
  type ActuationResult,
  type CapabilityManifest,
  type ContextRef,
  type ProvenancePayload,
  type ResolvedContext,
  type SseEvent,
  type Surface,
  type UnitDescriptor,
} from '@ge/contracts';
import { native, processNative, toContextNative } from '@ge/content';
import type { NativeContent } from '@ge/content';
import { contextValueToQueryPart, type QueryPart } from '@ge/gemini-client';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import type { DocBridge } from '@ge/runtime';

/**
 * INTERPLAY — content-grounding seam: @ge/content + @ge/gemini-client + @ge/runtime.
 *
 * Wires the REAL content pipeline (native Office object model → `Block[]` → budgeted `Chunk[]` →
 * `ResolvedContext[]`) into a REAL `AssistSession`, drives it through a scripted fake client that
 * replays `token` → `citation` → `provenance` → `done` SSE, and asserts the grounded answer, the
 * citations, and a provenance-stamped write all flow back across the seam. The ONLY mocked things
 * are the two outermost boundaries: the Google NETWORK (the scripted `StreamAssistClient`) and the
 * Office HOST (a recording `DocBridge` that captures actuations into in-memory "durable metadata").
 * Everything between — `processNative`/`chunkBlocks`/`toContextNative`, `SessionContext`,
 * `contextValueToQueryPart`, `AssistSession.ask`/`apply` — is the REAL implementation.
 */

// ---------------------------------------------------------------------------
// Outer boundary 1: the Google NETWORK — a capturing scripted StreamAssistClient.
// Mirrors the harness `scriptedClient` but ALSO captures the `context` (the query.parts that went
// on the wire) so the test can read back exactly what the content pipeline shipped to grounding.
// ---------------------------------------------------------------------------

interface ScriptedTurn {
  text: string;
  citations?: ProvenancePayload['sources'];
}

interface CapturingClient {
  client: StreamAssistClient;
  /** Every turn's outbound `context` (the resolved parts the session put on the wire). */
  sentContext: ResolvedContext[][];
  /** Every turn's outbound `context` mapped through the REAL `contextValueToQueryPart`. */
  sentParts: QueryPart[][];
}

function capturingClient(turns: ScriptedTurn[], provenance: ProvenancePayload): CapturingClient {
  const sentContext: ResolvedContext[][] = [];
  const sentParts: QueryPart[][] = [];
  let i = 0;

  const stream = async function* (
    _req: unknown,
    opts: StreamOptions = {},
  ): AsyncGenerator<SseEvent> {
    const ctx = opts.context ?? [];
    sentContext.push(ctx);
    // Exercise the REAL gemini-client mapping the session would use to build query.parts[].
    sentParts.push(ctx.map((c) => contextValueToQueryPart(c.value)));
    const turn = turns[i++] ?? { text: 'done' };
    yield { type: 'token', text: turn.text };
    for (const source of turn.citations ?? []) yield { type: 'citation', source };
    yield { type: 'provenance', payload: provenance };
    yield { type: 'done' };
  };

  return {
    client: { stream } as unknown as StreamAssistClient,
    sentContext,
    sentParts,
  };
}

// ---------------------------------------------------------------------------
// Outer boundary 2: the Office HOST — a recording DocBridge.
// `apply()` is REAL runtime code; this only fakes the host edge it actuates against, recording the
// landed request into an in-memory "durable metadata" store so the test can read provenance back.
// ---------------------------------------------------------------------------

interface RecordingBridge extends DocBridge {
  /** The actuation requests that actually landed (post-gate), as the host would have persisted. */
  readonly landed: ActuationRequest[];
}

function recordingBridge(surface: Surface): RecordingBridge {
  const landed: ActuationRequest[] = [];
  const manifest: CapabilityManifest = {
    surface,
    contextKinds: ['paragraph'],
    actuations: [
      { kind: 'tracked-change', surface, title: 'Insert tracked change', reversible: true },
    ],
  };
  return {
    surface,
    landed,
    getCapabilities: () => manifest,
    listContext: async (): Promise<ContextRef[]> => [],
    resolveContext: async (): Promise<ResolvedContext[]> => [],
    actuate: async (request: ActuationRequest): Promise<ActuationResult> => {
      // The host persists the request into durable metadata (here: push into `landed`).
      landed.push(request);
      return { ok: true, changeId: request.changeId, kind: request.kind, location: 'para:2' };
    },
  };
}

// ---------------------------------------------------------------------------
// Fixtures.
// ---------------------------------------------------------------------------

const PROVENANCE: ProvenancePayload = {
  agentId: 'gemini-enterprise:sim',
  identity: 'sim.user@acme',
  timestamp: '2026-06-24T00:00:00Z',
  sources: [{ title: 'Vendor Risk Notebook', uri: 'nb://vendor-risk' }],
  contentHash: 'sim-hash',
};

function bareUnit(surface: Surface): UnitDescriptor {
  return { connectors: [], surfaceContext: { kind: surface } };
}

/** A small native Word document with structure the chunker groups under its heading breadcrumb. */
function smallContractDoc(): NativeContent {
  return {
    sourceId: 'word:body',
    title: 'Master Services Agreement',
    surface: 'word',
    blocks: [
      native.heading('5. Service Levels', 1, 'cc:5'),
      native.paragraph('The SLA is 99.5% measured monthly.', 'cc:6'),
      native.paragraph('Credits accrue below the threshold.', 'cc:7'),
    ],
  };
}

async function drainAsk(
  session: AssistSession,
  query: string,
): Promise<{ answer: string; events: SseEvent[] }> {
  let answer = '';
  const events: SseEvent[] = [];
  for await (const ev of session.ask(query)) {
    events.push(ev);
    if (ev.type === 'token') answer += ev.text;
  }
  return { answer, events };
}

describe('content → gemini-client → runtime grounding interplay', () => {
  it('REAL native chunks ship as query.parts, the grounded answer + citations flow back, and an approved write is provenance-stamped into durable metadata', async () => {
    // 1. REAL content pipeline: native blocks → budgeted chunks → attach-ready ResolvedContext.
    const doc = smallContractDoc();
    const processed = processNative(doc); // REAL chunkBlocks
    const resolved = toContextNative(doc); // REAL chunk → ResolvedContext mapping
    // The pipeline grouped the two body paragraphs with their heading into one section chunk.
    expect(processed.chunks.length).toBeGreaterThan(0);
    expect(resolved.length).toBe(processed.chunks.length);
    // The chunk carried the section breadcrumb + the verbatim SLA claim — this is what grounds.
    const grounded = resolved.find((r) => r.value.as === 'text' && r.value.text.includes('99.5%'));
    expect(grounded).toBeDefined();

    // 2. REAL AssistSession over the recording bridge + capturing scripted client.
    const bridge = recordingBridge('word');
    const client = capturingClient(
      [
        {
          text: 'The agreement commits to a 99.5% monthly SLA [1].',
          citations: PROVENANCE.sources,
        },
      ],
      PROVENANCE,
    );
    const session = new AssistSession(bridge, client.client, { unit: bareUnit('word') });

    // Attach the REAL resolved chunks to the session's REAL SessionContext.
    for (const ctx of resolved) session.context.add(ctx);
    expect(session.context.size).toBe(resolved.length);

    // 3. Drive the REAL ask() — token + citation + provenance + done flow through the real session.
    const { answer, events } = await drainAsk(session, 'What is the SLA commitment?');

    // The grounded answer accumulated from the streamed tokens.
    expect(answer).toContain('99.5%');
    // Citations captured on the real session (the citation SSE flowed back through ask()).
    expect(session.sources).toHaveLength(PROVENANCE.sources.length);
    expect(session.sources[0]?.uri).toBe('nb://vendor-risk');
    expect(events.some((e) => e.type === 'provenance')).toBe(true);

    // CROSS-BOUNDARY: the content pipeline's chunk text actually went on the wire as query context.
    expect(client.sentContext).toHaveLength(1);
    const wireText = client.sentContext[0]!.map((c) =>
      c.value.as === 'text' ? c.value.text : '',
    ).join('\n');
    expect(wireText).toContain('99.5%');
    // And the REAL gemini-client mapping turned each chunk into a text query.part.
    expect(client.sentParts[0]!.every((p) => 'text' in p)).toBe(true);
    expect(client.sentParts[0]!.length).toBe(resolved.length);

    // 4. A provenance-stamped, reversible write flows back through the REAL apply() path.
    const result = await session.apply(
      'tracked-change',
      {
        text: 'The SLA is 99.5% measured monthly. [grounded]',
        target: { matchText: 'The SLA is 99.5% measured monthly.' },
      },
      asChangeId('chg-1'),
    );
    expect(result.ok).toBe(true);

    // OBSERVABLE: the write landed in the host's durable metadata carrying the stream's provenance.
    expect(bridge.landed).toHaveLength(1);
    const landed = bridge.landed[0]!;
    expect(landed.kind).toBe('tracked-change');
    expect(landed.surface).toBe('word');
    // The provenance the engine streamed (agent id, signed-in identity, sources) was stamped on.
    expect(landed.provenance?.agentId).toBe('gemini-enterprise:sim');
    expect(landed.provenance?.identity).toBe('sim.user@acme');
    expect(landed.provenance?.sources?.[0]?.uri).toBe('nb://vendor-risk');
  });

  it('a budget-boundary input is CHUNKED/TRUNCATED by the real pipeline, and the bounded parts (not one mega-block) ship on the wire', async () => {
    // One oversized paragraph (many sentences) that alone exceeds a small per-chunk token budget,
    // forcing the REAL chunker to recursively split it (paragraph → sentences) with overlap.
    const sentences = Array.from(
      { length: 40 },
      (_, i) => `Clause ${i + 1} states that the parties shall act in good faith at all times.`,
    ).join(' ');
    const doc: NativeContent = {
      sourceId: 'word:long',
      title: 'Long Clause',
      surface: 'word',
      blocks: [native.paragraph(sentences, 'cc:42')],
    };

    const MAX = 60; // small per-chunk budget to force the boundary split
    const processed = processNative(doc, { maxTokens: MAX, overlapTokens: 10 });
    const resolved = toContextNative(doc, { maxTokens: MAX, overlapTokens: 10 });

    // OBSERVABLE budget enforcement: the single over-budget block was split into MANY chunks…
    expect(processed.chunks.length).toBeGreaterThan(1);
    expect(resolved.length).toBe(processed.chunks.length);
    // …and every emitted chunk's own token estimate respects the boundary (none is a mega-block).
    for (const chunk of processed.chunks) {
      expect(chunk.meta.tokensEstimate).toBeLessThanOrEqual(MAX);
    }
    // The split preserved the native locator on every piece (write-back stays anchored).
    for (const chunk of processed.chunks) {
      expect(chunk.meta.anchor.locator).toBe('cc:42');
    }

    // Drive the REAL session: the bounded parts — not one mega-block — go on the wire.
    const bridge = recordingBridge('word');
    const client = capturingClient([{ text: 'Acknowledged.' }], PROVENANCE);
    const session = new AssistSession(bridge, client.client, { unit: bareUnit('word') });
    for (const ctx of resolved) session.context.add(ctx);

    await drainAsk(session, 'Summarize the clauses.');

    // CROSS-BOUNDARY: the wire carried the multiple bounded parts the pipeline truncated to…
    expect(client.sentParts).toHaveLength(1);
    expect(client.sentParts[0]!.length).toBe(resolved.length);
    expect(client.sentParts[0]!.length).toBeGreaterThan(1);
    // …and no single part on the wire blew the per-chunk budget (truncation actually held end-to-end).
    for (const part of client.sentParts[0]!) {
      expect('text' in part).toBe(true);
      if ('text' in part) {
        // estimateTokens is a heuristic; assert the part text is bounded by a generous char ceiling
        // derived from the token budget (4 chars/token Latin) so a mega-block would visibly fail.
        expect(part.text.length).toBeLessThanOrEqual(MAX * 8);
      }
    }
  });
});
