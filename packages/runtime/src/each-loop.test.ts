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
import type { CommandLoopEvent, PlanEffect } from './assist-session.js';
import type { DocBridge } from './bridge.js';

/**
 * C1 `each` through the REAL command loop: the block expands into concrete effects that share ONE
 * plan preview / approval, and an over-limit table is refused before anything is staged.
 */

class ExcelBridge implements DocBridge {
  readonly surface = 'excel' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'excel',
      contextKinds: ['range'],
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
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
    const gfm = '| cell | amount |\n| --- | --- |\n| B2 | 100 |\n| B3 | 250 |\n| B4 | 50 |';
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: { as: 'text', text: gfm, mimeType: 'text/markdown' },
      },
    ]);
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

function scriptedFetch(turns: string[], bodies: string[]): typeof fetch {
  let call = 0;
  const impl = vi.fn(async (_url: string, init?: { body?: string }) => {
    bodies.push(init?.body ?? '');
    const text = turns[Math.min(call, turns.length - 1)] ?? '```cmd\ndone\n```';
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(streamOf([JSON.stringify([chunk])]), { status: 200 });
  });
  return impl as unknown as typeof fetch;
}

async function collectLoop(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const out: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) out.push(e);
  return out;
}

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' };
const tokens = { getAccessToken: () => Promise.resolve('t') };
const unit = { connectors: [], surfaceContext: { kind: 'excel' as const } };

function session(bridge: ExcelBridge, turns: string[], bodies: string[] = []): AssistSession {
  const client = new StreamAssistClient(tokens, cfg, scriptedFetch(turns, bodies));
  return new AssistSession(bridge, client, { unit, context: { docState: false } });
}

const eachEvents = (events: Array<SseEvent | CommandLoopEvent>) =>
  events.filter((e) => e.type === 'each-expanded') as Array<
    Extract<CommandLoopEvent, { type: 'each-expanded' }>
  >;

describe('AssistSession.runCommands — C1 each loop', () => {
  it('a 3-row each expands into 3 concrete writes under ONE approval', async () => {
    const bridge = new ExcelBridge();
    const s = session(bridge, [
      '```cmd\nlet $t = read Sales!A1:B4\n```',
      '```cmd\neach $r in $t:\n  set Summary!$r.cell $r.amount\nend\n```',
      '```cmd\ndone\n```',
    ]);
    const previews: PlanEffect[][] = [];
    const events = await collectLoop(
      s.runCommands('copy the amounts', {
        approvePlan: (effects) => {
          previews.push(effects);
          return true;
        },
      }),
    );

    const [expanded] = eachEvents(events);
    expect(expanded).toMatchObject({
      rows: 3,
      lines: ['set Summary!B2 100', 'set Summary!B3 250', 'set Summary!B4 50'],
    });
    // One plan, three concrete effects — the approval card saw every substituted write.
    expect(previews).toHaveLength(1);
    expect(previews[0]!.map((e) => e.command)).toEqual([
      'set Summary!B2 100',
      'set Summary!B3 250',
      'set Summary!B4 50',
    ]);
    expect(bridge.applied.map((r) => r.params.cells)).toEqual([[['100']], [['250']], [['50']]]);
    expect(events.filter((e) => e.type === 'write-result')).toHaveLength(3);
  });

  it('a 51-row table is refused before approval — nothing previewed, nothing applied', async () => {
    const bridge = new ExcelBridge();
    const rows = Array.from({ length: 51 }, (_, i) => `["B${i + 2}","${i}"]`).join(',');
    const bodies: string[] = [];
    const s = session(
      bridge,
      [
        `\`\`\`cmd\nlet $t = table [["cell","amount"],${rows}]\n\`\`\``,
        '```cmd\neach $r in $t:\n  set Summary!$r.cell $r.amount\nend\n```',
        '```cmd\ndone\n```',
      ],
      bodies,
    );
    const approve = vi.fn(() => true);
    const events = await collectLoop(s.runCommands('copy the amounts', { approvePlan: approve }));

    expect(eachEvents(events)[0]).toMatchObject({ rows: 0, lines: [] });
    expect(approve).not.toHaveBeenCalled();
    expect(bridge.applied).toHaveLength(0);
    expect(events.some((e) => e.type === 'plan-preview')).toBe(false);
    // The corrective is fed back in the next turn's result block, naming the limit so the model
    // narrows the table instead of retrying.
    expect(bodies[2]).toContain('the loop limit is 50');
  });

  it('an expansion over the write cap is refused as a whole', async () => {
    const bridge = new ExcelBridge();
    const bodies: string[] = [];
    const s = session(
      bridge,
      [
        '```cmd\nlet $t = read Sales!A1:B4\n```',
        '```cmd\neach $r in $t:\n  set Summary!$r.cell $r.amount\n  set Other!$r.cell $r.amount\nend\n```',
        '```cmd\ndone\n```',
      ],
      bodies,
    );
    const approve = vi.fn(() => true);
    const events = await collectLoop(
      s.runCommands('copy twice', { approvePlan: approve, maxWritesPerTurn: 4 }),
    );
    expect(eachEvents(events)[0]).toMatchObject({ rows: 0, lines: [] });
    expect(approve).not.toHaveBeenCalled();
    expect(bridge.applied).toHaveLength(0);
    expect(bodies[2]).toContain('write cap of 4');
  });
});
