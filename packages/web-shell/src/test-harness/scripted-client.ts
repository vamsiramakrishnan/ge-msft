/**
 * A scripted fake `StreamAssistClient` for full-stack integration tests. The REAL
 * {@link "@ge/runtime"!AssistSession} only ever calls `.stream(req, opts)`, so a fake that replays a
 * programmed sequence of model turns — each a string of answer text, typically carrying a ```cmd
 * block — drives the real command loop (`runCommands`) end to end without a network or Google
 * credential. Mirrors the runtime's own `fakeClient` test helper, lifted here so the web-shell
 * integration tests share one scripted-transcript contract.
 *
 * Each turn is wrapped as `token` → `provenance` → `done` SSE events, so the controller's streamed
 * assistant message, citations, and provenance-stamped writes all exercise their real paths. The
 * provenance identity is seeded so an approved write records a real provenance record into the host.
 */

import type { AssistRequest, SseEvent, ProvenancePayload } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';

/** The default provenance payload stamped on each scripted turn (the signed-in user). */
const DEFAULT_PROVENANCE: ProvenancePayload = {
  agentId: 'gemini-enterprise:sim',
  identity: 'sim.user@acme',
  timestamp: '2026-06-23T00:00:00Z',
  sources: [],
  contentHash: 'sim-hash',
  sessionId: 'sess_sim' as never,
};

/** A scripted turn: the raw answer text (with an optional ```cmd block) and optional citations. */
export interface ScriptedTurn {
  text: string;
  /** Citations to emit as `citation` SSE events before the answer settles. */
  citations?: ProvenancePayload['sources'];
}

/** A scripted fake client + the queries it received (for asserting what was sent to the model). */
export interface ScriptedClient {
  client: StreamAssistClient;
  /** Every `req.query` (flattened to text) the loop sent, in order. */
  queries: string[];
}

/** Flatten an `AssistRequest.query` (string or multi-part) to a single string for assertions. */
function queryText(req: AssistRequest): string {
  const q: unknown = (req as { query?: unknown }).query;
  if (typeof q === 'string') return q;
  if (q && typeof q === 'object' && 'parts' in q) {
    const parts = (q as { parts?: Array<{ text?: string }> }).parts ?? [];
    return parts.map((p) => p.text ?? '').join('\n');
  }
  return '';
}

/**
 * Build a scripted client from a list of turns. Each `runCommands`/`ask` turn pulls the next entry;
 * once the script is exhausted it replays a terminal `done` block so a runaway loop still ends.
 */
export function scriptedClient(
  turns: ReadonlyArray<string | ScriptedTurn>,
  provenance: ProvenancePayload = DEFAULT_PROVENANCE,
): ScriptedClient {
  const queries: string[] = [];
  let i = 0;
  const normalized = turns.map((t) => (typeof t === 'string' ? { text: t } : t));

  const stream = async function* (
    req: AssistRequest,
    _opts: StreamOptions,
  ): AsyncGenerator<SseEvent> {
    queries.push(queryText(req));
    const turn = normalized[i++] ?? { text: '```cmd\ndone\n```' };
    yield { type: 'token', text: turn.text };
    for (const source of turn.citations ?? []) {
      yield { type: 'citation', source };
    }
    yield { type: 'provenance', payload: provenance };
    yield { type: 'done' };
  };

  const client = { stream } as unknown as StreamAssistClient;
  return { client, queries };
}
