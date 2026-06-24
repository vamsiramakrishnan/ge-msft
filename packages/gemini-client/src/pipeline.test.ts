import { describe, it, expect } from 'vitest';
import { toContext } from '@ge/content';
import type { AssistRequest } from '@ge/contracts';
import { SessionContext } from './session-context.js';
import { buildStreamAssistRequest } from './stream-assist.js';

/**
 * Interplay test: @ge/content processes raw host content into ResolvedContext, which
 * SessionContext attaches and the request builder turns into a multi-part streamAssist
 * query. This is the full content → context → query.parts[] path the bridges drive.
 */
describe('content → SessionContext → streamAssist query.parts', () => {
  const DOC = `# Service Levels

## Availability
The services are available 99.5% of the time.

## Support
P1 incidents get a 1 hour response.`;

  const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'u@acme' };
  const req: AssistRequest = {
    intent: 'ask',
    query: 'Is the SLA below our floor?',
    unit: { connectors: [], surfaceContext: { kind: 'word' } },
  };

  it('attaches processed chunks as contextualized markdown parts, question last', () => {
    const sc = new SessionContext();
    for (const c of toContext(
      { sourceId: 'word:body', text: DOC, format: 'markdown', title: 'MSA', surface: 'word' },
      { maxTokens: 50 },
    )) {
      sc.add(c);
    }
    expect(sc.size).toBeGreaterThan(1);

    const body = buildStreamAssistRequest(req, cfg, undefined, sc.list());
    const parts = (body.query as { parts: { text?: string }[] }).parts;

    // Each context part is contextualized markdown; the user's question is the final part.
    expect(parts[0]!.text).toMatch(/^\[MSA/);
    expect(parts.some((p) => p.text?.includes('99.5%'))).toBe(true);
    expect(parts.at(-1)!.text).toBe('Is the SLA below our floor?');
  });

  it('carries write-back anchors through to the attached context', () => {
    const ctx = toContext(
      { sourceId: 'word:body', text: DOC, format: 'markdown', surface: 'word' },
      { maxTokens: 50 },
    );
    const avail = ctx.find((c) => c.ref.preview?.includes('99.5%'))!;
    expect(avail.ref.anchor?.contextHint).toContain('Availability');
    expect(avail.ref.anchor?.locator).toMatch(/^chars:/);
  });
});
