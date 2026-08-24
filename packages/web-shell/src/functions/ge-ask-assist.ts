import type { AssistRequest, SseEvent } from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import type { GeAskAssist } from './ge-ask.js';

/**
 * Yield just the token text from a contracts SSE event stream. Errors surface as thrown
 * `Error`s (the cell shows the message); `done`/`provenance`/`citation` are irrelevant to a
 * single-cell answer and are dropped here rather than silently flattened into text.
 */
export async function* tokensFromSse(events: AsyncIterable<SseEvent>): AsyncIterable<string> {
  for await (const ev of events) {
    if (ev.type === 'token') yield ev.text;
    if (ev.type === 'error') throw new Error(ev.message);
  }
}

/**
 * Bind the function's assist seam to a client-direct `StreamAssistClient`: one `ask` turn,
 * grounded on nothing but the framed workbook data in the query itself. The engine owns Model
 * Armor and grounding; this adapter owns only transport shaping.
 */
export function createGeAskAssist(client: StreamAssistClient): GeAskAssist {
  return ({ text, signal }) => {
    const req: AssistRequest = {
      intent: 'ask',
      unit: { connectors: [], surfaceContext: { kind: 'excel' } },
      query: text,
    };
    return tokensFromSse(client.stream(req, signal ? { signal } : {}));
  };
}
