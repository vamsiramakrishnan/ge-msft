import { describe, it, expect } from 'vitest';
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
import type { CommandLoopEvent } from './assist-session.js';
import type { DocBridge } from './bridge.js';

/**
 * ADR-0007 / B3 — undo via recorded inverses. A landed receipt that carries a usable `inverse`
 * on a bridge implementing `applyInverse` is offered for undo (`canUndo`); the pane's undo action
 * (`runAnalysis({ kind: 'undo' })`) previews an undo effect, takes the plan approval, and dispatches
 * the bridge's inverse applier — never `actuate` — with a receipt naming the reversed change.
 */

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' };
const tokens = { getAccessToken: () => Promise.resolve('t') };
const unit = { connectors: [], surfaceContext: { kind: 'excel' as const } };

function scriptedFetch(turns: string[]): typeof fetch {
  let call = 0;
  return (async () => {
    const text = turns[Math.min(call, turns.length - 1)] ?? '```cmd\ndone\n```';
    call += 1;
    const chunk = {
      sessionInfo: { session: 'sess_1' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    return new Response(JSON.stringify([chunk]), { status: 200 });
  }) as unknown as typeof fetch;
}

async function collect(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const out: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) out.push(e);
  return out;
}

class InverseBridge implements DocBridge {
  readonly surface = 'excel' as const;
  actuated: ActuationRequest[] = [];
  inverted: ActuationRequest[] = [];
  /** Present only when the fixture is built with an applier (a prototype method could not be absent). */
  applyInverse?: (request: ActuationRequest) => Promise<ActuationResult>;
  constructor(
    private readonly inverse: NonNullable<ActuationResult['inverse']>,
    withApplier = true,
  ) {
    if (withApplier)
      this.applyInverse = (request) => {
        this.inverted.push(request);
        return Promise.resolve({
          ok: true,
          changeId: request.changeId,
          kind: request.kind,
          location: 'Table1',
          verification: { status: 'verified' },
        });
      };
  }
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'excel',
      contextKinds: [],
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'set', reversible: true }],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.actuated.push(request);
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'Table1',
      inverse: this.inverse,
      verification: { status: 'verified' },
    });
  }
}

const tableInverse = { op: 'delete-object' as const, objectType: 'table' as const, name: 'Table1' };

async function landOne(bridge: InverseBridge): Promise<{ session: AssistSession; id: string }> {
  const client = new StreamAssistClient(
    tokens,
    cfg,
    scriptedFetch(['```cmd\nset Sales!A1 5\n```', '```cmd\ndone\n```']),
  );
  const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
  const events = await collect(session.runCommands('fill A1', { approvePlan: () => true }));
  const landed = events.find((e) => e.type === 'write-result') as Extract<
    CommandLoopEvent,
    { type: 'write-result' }
  >;
  expect(landed?.result.ok).toBe(true);
  return { session, id: landed.changeId };
}

describe('AssistSession undo via recorded inverses (B3)', () => {
  it('canUndo is true for a landed receipt with a usable inverse on a bridge that applies inverses', async () => {
    const { session, id } = await landOne(new InverseBridge(tableInverse));
    expect(session.canUndo(id)).toBe(true);
    expect(session.undoableChanges()).toEqual([id]);
    expect(session.canUndo('never-landed')).toBe(false);
  });

  it('canUndo is false without applyInverse, and for a not-reversible marker', async () => {
    const noApplier = await landOne(new InverseBridge(tableInverse, false));
    expect(noApplier.session.canUndo(noApplier.id)).toBe(false);
    const marked = await landOne(new InverseBridge({ op: 'not-reversible', reason: 'no handle' }));
    expect(marked.session.canUndo(marked.id)).toBe(false);
  });

  it('runAnalysis undo previews an undo effect, dispatches applyInverse after approval, and names the reversed change', async () => {
    const bridge = new InverseBridge(tableInverse);
    const { session, id } = await landOne(bridge);
    const approved: Array<{ command: string; undoOf?: string }> = [];
    const events = await collect(
      session.runAnalysis(
        { kind: 'undo', id },
        {
          approvePlan: (effects) => {
            approved.push(
              ...effects.map((e) => ({ command: e.command, undoOf: e.request.undo?.of })),
            );
            return true;
          },
        },
      ),
    );
    const preview = events.find((e) => e.type === 'plan-preview') as Extract<
      CommandLoopEvent,
      { type: 'plan-preview' }
    >;
    expect(preview.effects[0]?.command).toBe(`undo ${id}`);
    expect(preview.effects[0]?.request.undo).toEqual({ of: id, inverse: tableInverse });
    expect(preview.effects[0]?.request.kind).toBe('write-cells');
    expect(preview.effects[0]?.reversible).toBe(false);
    expect(preview.effects[0]?.dryRun?.resolved).toContain('remove the table');
    expect(approved).toEqual([{ command: `undo ${id}`, undoOf: id }]);
    // The inverse applier ran once; `actuate` saw only the original write.
    expect(bridge.inverted).toHaveLength(1);
    expect(bridge.inverted[0]?.undo?.of).toBe(id);
    expect(bridge.actuated).toHaveLength(1);
    const result = events.find((e) => e.type === 'write-result') as Extract<
      CommandLoopEvent,
      { type: 'write-result' }
    >;
    expect(result.result).toMatchObject({ ok: true, undoOf: id });
    // The reversed change is no longer offered for undo, and the undo receipt itself never is.
    expect(session.canUndo(id)).toBe(false);
    expect(session.canUndo(result.changeId)).toBe(false);
  });

  it('a declined approval dispatches nothing and keeps the change undoable', async () => {
    const bridge = new InverseBridge(tableInverse);
    const { session, id } = await landOne(bridge);
    const events = await collect(
      session.runAnalysis({ kind: 'undo', id }, { approvePlan: () => false }),
    );
    const result = events.find((e) => e.type === 'write-result') as Extract<
      CommandLoopEvent,
      { type: 'write-result' }
    >;
    expect(result.result.error?.code).toBe('plan_unapproved');
    expect(bridge.inverted).toHaveLength(0);
    expect(session.canUndo(id)).toBe(true);
  });
});
