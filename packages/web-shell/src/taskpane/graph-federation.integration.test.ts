import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  EstateRefSchema,
  ResolvedContextSchema,
  type EstateRef,
  type ResolvedContext,
  type AssistRequest,
  type SseEvent,
  type ProvenancePayload,
} from '@ge/contracts';
import { GraphClient, type GraphTokenSource } from '@ge/graph-client';
import { GRAPH_SCOPES } from '@ge/graph-client';
import { AssistSession } from '@ge/runtime';
import {
  contextValueToQueryPart,
  type StreamAssistClient,
  type StreamOptions,
} from '@ge/gemini-client';
import { installFakeWord, type WordSimulator } from '../test-harness/index.js';
import { selectBridge } from './select-bridge.js';

/**
 * INTERPLAY — `graph-federation`. Wires the REAL Graph reader (`@ge/graph-client`) resolving
 * federated SharePoint/OneDrive sources into a REAL research unit / `AssistSession` context
 * (`@ge/runtime`), across the REAL contracts grammar (`@ge/contracts`).
 *
 * What is REAL (everything between the boundaries):
 *   - the REAL `GraphClient` — its `/search/query` + `/drives/.../items/...` transport, its
 *     401-retry, its Microsoft-Search-hit → `EstateRef` mapping and `EstateRef` → `ResolvedContext`
 *     resolution (`driveItemToContext`, the reference-over-inline policy);
 *   - the REAL `AssistSession` + its `SessionContext` (`@ge/gemini-client`) — the research unit's
 *     live context set that becomes the Discovery Engine `query.parts[]` on the wire;
 *   - the REAL `ContextModel` (the event-fed working brief) reached via `session.model`;
 *   - the REAL Word `DocBridge` (`selectBridge('word')`), so the session has a real surface.
 *
 * What is FAKED (only the two outer boundaries):
 *   - the Graph NETWORK — a `vi.fn` fetch returning Graph/Search JSON payloads;
 *   - the Google/model NETWORK — an observing fake `StreamAssistClient` that records the
 *     `context` (query.parts) the session puts on the wire;
 *   - the Office HOST — `installFakeWord()` sets `globalThis.Word`/`globalThis.Office`.
 *
 * The assertions are cross-boundary + observable: the Graph-resolved SharePoint/OneDrive source
 * actually lands in the context the session sends; the delegated scope/token actually reached the
 * Graph fetch (delegated scoping is respected, not assumed); and a 429/throttle is SURFACED
 * (propagated to the caller), never silently swallowed into an empty context.
 */

const PROVENANCE: ProvenancePayload = {
  agentId: 'gemini-enterprise:sim',
  identity: 'sim.user@acme',
  timestamp: '2026-06-24T00:00:00Z',
  sources: [],
  contentHash: 'sim-hash',
  sessionId: 'sess_graph' as never,
};

function jsonResponse(obj: unknown, status = 200): Response {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/** A Graph token source whose token + invalidation are observable (the delegated NAA token). */
function delegatedTokens(token = 'delegated-graph-token'): GraphTokenSource & {
  invalidate: ReturnType<typeof vi.fn>;
  scopeCalls: string[][];
} {
  const scopeCalls: string[][] = [];
  return {
    getGraphToken: (scopes: string[]) => {
      scopeCalls.push(scopes);
      return Promise.resolve(token);
    },
    invalidate: vi.fn(),
    scopeCalls,
  };
}

/**
 * A boundary fake for the model NETWORK. The REAL `AssistSession.ask` calls `.stream(req, opts)`
 * exactly once per turn, passing `opts.context` = the session's resident context set (the research
 * unit's query.parts). We record both so a test can assert what the federated sources became on the
 * wire — without a network or a Google credential.
 */
function observingClient(): {
  client: StreamAssistClient;
  contexts: ResolvedContext[][];
  requests: AssistRequest[];
} {
  const contexts: ResolvedContext[][] = [];
  const requests: AssistRequest[] = [];
  const stream = async function* (
    req: AssistRequest,
    opts: StreamOptions,
  ): AsyncGenerator<SseEvent> {
    requests.push(req);
    contexts.push([...((opts.context as ResolvedContext[] | undefined) ?? [])]);
    yield { type: 'token', text: 'ack' };
    yield { type: 'provenance', payload: PROVENANCE };
    yield { type: 'done' };
  };
  return { client: { stream } as unknown as StreamAssistClient, contexts, requests };
}

async function drain(gen: AsyncGenerator<SseEvent>): Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  for await (const _ of gen) {
    /* consume */
  }
}

let sim: WordSimulator | undefined;
afterEach(() => {
  sim?.restore();
  sim = undefined;
  vi.restoreAllMocks();
});

describe('graph-federation interplay (graph-client + runtime + contracts)', () => {
  it('resolves a federated SharePoint source via REAL Graph search and lands it in the session context on the wire', async () => {
    sim = installFakeWord();
    const tokens = delegatedTokens();

    // The Graph NETWORK: a single fetch that answers /search/query then the drive-item read.
    const fetchImpl = vi.fn(async (url: string, _init?: RequestInit) => {
      if (url.includes('/search/query')) {
        return jsonResponse({
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      hitId: 'h-msa',
                      summary: 'Master Services Agreement — DealRoom',
                      resource: {
                        '@odata.type': '#microsoft.graph.driveItem',
                        id: 'd-msa',
                        name: 'MSA.docx',
                        webUrl: 'https://contoso.sharepoint.com/sites/DealRoom/MSA.docx',
                        parentReference: { driveId: 'drv-dealroom' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        });
      }
      // resolveEstateRef('drive-item') → GET /drives/{driveId}/items/{id}
      return jsonResponse({
        id: 'd-msa',
        name: 'MSA.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/DealRoom/MSA.docx',
        size: 48213,
        parentReference: { driveId: 'drv-dealroom' },
      });
    });

    const graph = new GraphClient(tokens, {}, fetchImpl as never);

    // --- REAL Graph: search the estate as the user, then resolve the federated hit to context. ---
    const refs = await graph.search('vendor risk MSA', ['driveItem']);
    expect(refs).toHaveLength(1);
    for (const r of refs) expect(() => EstateRefSchema.parse(r)).not.toThrow();
    const ref: EstateRef = refs[0]!;
    expect(ref).toMatchObject({ source: 'drive-item', id: 'd-msa', driveId: 'drv-dealroom' });

    const resolved = await graph.resolveEstateRef(ref);
    expect(resolved.length).toBeGreaterThan(0);
    for (const c of resolved) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    // Reference-over-inline: a SharePoint file grounds as a drive-document handle (not inlined binary).
    expect(resolved[0]!.value.as).toBe('drive-document');

    // --- REAL runtime: build the research unit / session over the REAL Word bridge, attach the
    // Graph-resolved federated source, and run a turn through the observing model boundary. ---
    const bridge = selectBridge('word');
    if (!bridge) throw new Error('no word bridge');
    const model = observingClient();
    const session = new AssistSession(bridge, model.client, {
      unit: {
        connectors: [{ type: 'sharepoint', mode: 'federated', scope: 'sites/DealRoom' }],
        surfaceContext: { kind: 'word' },
      },
    });
    for (const c of resolved) session.context.add(c);

    await drain(session.ask('summarise the SLA exposure in the MSA'));

    // CROSS-BOUNDARY assertion: the federated SharePoint source the REAL Graph reader resolved is
    // present in the context the REAL session put on the wire — Plane B converged into the unit.
    expect(model.contexts).toHaveLength(1);
    const onWire = model.contexts[0]!;
    const driveDoc = onWire.find((c) => c.value.as === 'drive-document');
    expect(driveDoc).toBeDefined();
    expect((driveDoc!.value as { driveId: string }).driveId).toBe('drv-dealroom');
    expect((driveDoc!.value as { title?: string }).title).toBe('MSA.docx');
    expect(driveDoc!.ref.id).toContain('graph:drive:');
    // ...and it converts (REAL gemini-client) into the Discovery Engine `driveDocumentReference`
    // wire part — a federated SharePoint handle, not inlined binary — closing the seam to the wire.
    const wirePart = contextValueToQueryPart(driveDoc!.value);
    expect(wirePart).toMatchObject({
      driveDocumentReference: { driveId: 'drv-dealroom', displayTitle: 'MSA.docx' },
    });
    // The request carried the federated connector unit (the research unit) too.
    expect(model.requests[0]!.unit.connectors[0]).toMatchObject({
      type: 'sharepoint',
      mode: 'federated',
      scope: 'sites/DealRoom',
    });
  });

  it('respects delegated scoping: the Graph reads carry the delegated files/search scopes + the user bearer token', async () => {
    sim = installFakeWord();
    const tokens = delegatedTokens('user-naa-token');
    const fetchImpl = vi.fn(async (url: string, _init?: RequestInit) => {
      if (url.includes('/search/query')) {
        return jsonResponse({
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      hitId: 'h1',
                      resource: {
                        '@odata.type': '#microsoft.graph.driveItem',
                        id: 'od-1',
                        name: 'Notes.docx',
                        parentReference: { driveId: 'drv-me' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        });
      }
      return jsonResponse({
        id: 'od-1',
        name: 'Notes.docx',
        parentReference: { driveId: 'drv-me' },
      });
    });

    const graph = new GraphClient(tokens, {}, fetchImpl as never);
    const refs = await graph.search('quarterly notes', ['driveItem']);
    await graph.resolveEstateRef(refs[0]!);

    // The search call requested the delegated SEARCH scope set (Files/Sites/Mail .Read), and the
    // drive read requested the delegated FILES scope set — never an app-wide / write scope.
    expect(tokens.scopeCalls[0]).toEqual([...GRAPH_SCOPES.search]);
    expect(tokens.scopeCalls[1]).toEqual([...GRAPH_SCOPES.files]);
    expect([...GRAPH_SCOPES.files]).not.toContain('Sites.ReadWrite.All');
    expect([...GRAPH_SCOPES.files]).not.toContain('Files.ReadWrite.All');

    // Every Graph call carried the signed-in user's delegated bearer token (scoped to the user,
    // end to end) — not a service principal.
    for (const call of fetchImpl.mock.calls) {
      const init = call[1] as RequestInit;
      expect(init.headers).toMatchObject({ Authorization: 'Bearer user-naa-token' });
    }
  });

  it('surfaces a 429 throttle from Graph (propagated to the caller) rather than silently degrading the unit to empty context', async () => {
    sim = installFakeWord();
    const tokens = delegatedTokens();
    // Graph throttles the resolve with 429 — and keeps throttling (no 401-style retry path covers it).
    const fetchImpl = vi.fn(async (url: string) => {
      if (url.includes('/search/query')) {
        return jsonResponse({
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      hitId: 'h1',
                      resource: {
                        '@odata.type': '#microsoft.graph.driveItem',
                        id: 'd-throttled',
                        name: 'Big.docx',
                        parentReference: { driveId: 'drv-x' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        });
      }
      return new Response('throttled', {
        status: 429,
        headers: { 'Retry-After': '30' },
      });
    });

    const graph = new GraphClient(tokens, {}, fetchImpl as never);
    const refs = await graph.search('big doc', ['driveItem']);
    expect(refs).toHaveLength(1);

    // The throttle is SURFACED — resolveEstateRef rejects with the status, not swallowed to [].
    await expect(graph.resolveEstateRef(refs[0]!)).rejects.toThrow(/429/);
    // It is NOT a 401, so the token was not invalidated (no spurious re-auth on a throttle).
    expect(tokens.invalidate).not.toHaveBeenCalled();

    // The unit/session therefore never silently attaches a phantom federated source on a throttle:
    // because the caller handled (surfaced) the error, the session context stays empty and the turn
    // still streams without a broken/empty grounding handle.
    const bridge = selectBridge('word');
    if (!bridge) throw new Error('no word bridge');
    const model = observingClient();
    const session = new AssistSession(bridge, model.client, {
      unit: { connectors: [], surfaceContext: { kind: 'word' } },
    });

    let attachError: unknown;
    try {
      const r = await graph.resolveEstateRef(refs[0]!);
      for (const c of r) session.context.add(c);
    } catch (err) {
      attachError = err;
    }
    expect(attachError).toBeDefined();
    expect(session.context.size).toBe(0);

    await drain(session.ask('what is in the MSA?'));
    expect(model.contexts[0]!.find((c) => c.value.as === 'drive-document')).toBeUndefined();
  });

  it('a resolved federated source also feeds the REAL ContextModel brief (estate-changed) and rides the next turn', async () => {
    sim = installFakeWord();
    const tokens = delegatedTokens();
    const fetchImpl = vi.fn(async () =>
      jsonResponse({
        id: 'd-1',
        name: 'Policy.pdf',
        webUrl: 'https://contoso.sharepoint.com/sites/Risk/Policy.pdf',
        parentReference: { driveId: 'drv-risk' },
      }),
    );
    const graph = new GraphClient(tokens, {}, fetchImpl as never);

    const bridge = selectBridge('word');
    if (!bridge) throw new Error('no word bridge');
    const model = observingClient();
    const session = new AssistSession(bridge, model.client, {
      unit: {
        connectors: [{ type: 'sharepoint', mode: 'federated', scope: 'sites/Risk' }],
        surfaceContext: { kind: 'word' },
      },
    });

    // The REAL graph resolution attaches the grounding handle...
    const resolved = await graph.resolveEstateRef({
      source: 'drive-item',
      id: 'd-1',
      driveId: 'drv-risk',
    });
    for (const c of resolved) session.context.add(c);

    // ...and the REAL ContextModel records the estate change as a working-brief delta (fold).
    const hint = session.model.observe({
      type: 'estate-changed',
      source: 'site',
      id: 'd-1',
    });
    expect(hint.commit).toBe('fold');
    expect(session.model.hasPending).toBe(true);

    await drain(session.ask('check the policy'));

    const onWire = model.contexts[0]!;
    // Both planes converged on the wire: the drive-document grounding handle AND the folded
    // working-context brief that narrates the estate change.
    expect(onWire.find((c) => c.value.as === 'drive-document')).toBeDefined();
    const brief = onWire.find((c) => c.ref.kind === 'brief');
    expect(brief).toBeDefined();
    expect((brief!.value as { text: string }).text).toContain('Estate item changed: site d-1');
    // The brief was committed (resident) once the turn fully streamed — not re-sent next turn.
    expect(session.model.hasPending).toBe(false);
  });
});
