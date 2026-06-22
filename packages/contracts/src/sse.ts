import { z } from 'zod';
import { FindingSchema, SourceRefSchema } from './finding.js';
import { ProvenancePayloadSchema } from './provenance.js';

/**
 * The streaming protocol. Streaming endpoints respond with `text/event-stream`;
 * the wire format is `event: <type>\ndata: <json>\n\n`.
 */
export const SseEventSchema = z.discriminatedUnion('type', [
  z.object({ type: z.literal('token'), text: z.string() }),
  z.object({ type: z.literal('finding'), finding: FindingSchema }),
  z.object({
    type: z.literal('slide'),
    title: z.string(),
    bullets: z.array(z.string()),
    sources: z.array(SourceRefSchema),
  }),
  z.object({ type: z.literal('citation'), source: SourceRefSchema }),
  // A grounded claim span. `start`/`end` are JS string (UTF-16) character indices
  // into the accumulated answer text, so a UI can `answerText.slice(start, end)`.
  z.object({
    type: z.literal('grounding-support'),
    start: z.number().int().nonnegative(),
    end: z.number().int().nonnegative(),
    score: z.number().optional(),
    sources: z.array(SourceRefSchema),
  }),
  // Outcome of a customer policy check (Model Armor / banned phrases).
  z.object({
    type: z.literal('policy'),
    verdict: z.enum(['allow', 'block']),
    reason: z.string().optional(),
  }),
  // Suggested follow-up questions emitted at the end of a turn.
  z.object({ type: z.literal('related-questions'), questions: z.array(z.string()) }),
  z.object({ type: z.literal('provenance'), payload: ProvenancePayloadSchema }), // sent before 'done'
  z.object({ type: z.literal('error'), code: z.string(), message: z.string() }),
  z.object({ type: z.literal('done') }),
]);
export type SseEvent = z.infer<typeof SseEventSchema>;

/** Serialize an event to the SSE wire format. */
export function serializeSseEvent(event: SseEvent): string {
  return `event: ${event.type}\ndata: ${JSON.stringify(event)}\n\n`;
}

/**
 * Parse one SSE `event:`/`data:` block back into a validated event.
 * Throws (via Zod) if the payload doesn't match the contract.
 */
export function parseSseEvent(block: string): SseEvent {
  let dataLine = '';
  for (const raw of block.split('\n')) {
    const line = raw.trimEnd();
    if (line.startsWith('data:')) {
      dataLine += line.slice(line.startsWith('data: ') ? 6 : 5);
    }
  }
  if (!dataLine) {
    throw new Error('parseSseEvent: no data line in block');
  }
  return SseEventSchema.parse(JSON.parse(dataLine));
}
