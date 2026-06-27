import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  Surface,
  SseEvent,
} from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import type { DocStateSnapshot } from '@ge/contracts';
import { AssistSession } from './assist-session.js';
import type { CommandLoopEvent } from './assist-session.js';
import type { DocBridge } from './bridge.js';

/* ──────────────────────────────── harness ──────────────────────────────── */

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' };
const tokens = { getAccessToken: () => Promise.resolve('t') };
const unit = { connectors: [], surfaceContext: { kind: 'word' as const } };

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

/**
 * A streamAssist fetch scripted per turn. Each element is the model's reply text for that turn;
 * once exhausted it keeps emitting `done`. Captures the bodies so we can assert what went on the
 * wire (e.g. the fed-back ```result``` block + a fresh <doc_state>).
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

const writeResults = (
  events: Array<SseEvent | CommandLoopEvent>,
): Array<Extract<CommandLoopEvent, { type: 'write-result' }>> =>
  events.filter((e) => e.type === 'write-result') as Array<
    Extract<CommandLoopEvent, { type: 'write-result' }>
  >;

const readResults = (
  events: Array<SseEvent | CommandLoopEvent>,
): Array<Extract<CommandLoopEvent, { type: 'read-result' }>> =>
  events.filter((e) => e.type === 'read-result') as Array<
    Extract<CommandLoopEvent, { type: 'read-result' }>
  >;

/**
 * A configurable bridge whose surface, capabilities, and optional read/doc-state ports are all
 * injectable. Lets one harness drive surface-context branches, read-intent branches, and effect
 * type-checks.
 */
class FlexBridge implements DocBridge {
  applied: ActuationRequest[] = [];
  actuateError: Error | null = null;
  constructor(
    readonly surface: Surface,
    private readonly opts: {
      actuations?: CapabilityManifest['actuations'];
      contextKinds?: CapabilityManifest['contextKinds'];
      // optional read ports
      captureDocState?: (() => Promise<DocStateSnapshot | undefined>) | null;
      readRange?: ((a1: string) => Promise<ResolvedContext[]>) | null;
      searchDocument?: ((q: string) => Promise<ResolvedContext[]>) | null;
    } = {},
  ) {
    if (opts.captureDocState !== null && opts.captureDocState !== undefined) {
      (this as { captureDocState?: unknown }).captureDocState = opts.captureDocState;
    }
    if (opts.readRange !== null && opts.readRange !== undefined) {
      (this as { readRange?: unknown }).readRange = opts.readRange;
    }
    if (opts.searchDocument !== null && opts.searchDocument !== undefined) {
      (this as { searchDocument?: unknown }).searchDocument = opts.searchDocument;
    }
  }
  getCapabilities(): CapabilityManifest {
    return {
      surface: this.surface,
      contextKinds: this.opts.contextKinds ?? [],
      actuations: this.opts.actuations ?? [],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    if (this.actuateError) return Promise.reject(this.actuateError);
    this.applied.push(request);
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
}

function gfmTable(): ResolvedContext[] {
  const gfm = '| region | amount |\n| --- | --- |\n| East | 100 |\n| West | 250 |';
  return [
    {
      ref: { id: 'xl:A1', kind: 'range', surface: 'excel', title: 'A1', live: false },
      value: { as: 'text', text: gfm, mimeType: 'text/markdown' },
    },
  ];
}

function snapshot(surface: Surface = 'word'): DocStateSnapshot {
  return {
    surface,
    version: 1,
    capturedAt: '2026-06-22T00:00:00.000Z',
    title: 'Doc',
    outline: [{ level: 1, text: 'Heading One' }],
    inventory: [],
  };
}

/* ───────────────────── runReadIntent — every read branch ─────────────────── */

describe('AssistSession.runCommands — read intents (ADR-0003 dispatch branches)', () => {
  it('outline dispatches to captureDocState → rendered doc_state', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    const { fetch } = scriptedFetch(['```cmd\noutline\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('show outline'));
    const r = readResults(events)[0];
    expect(r?.intentLabel).toBe('outline');
    expect(String(r?.result)).toContain('Heading One');
  });

  it('outline with no captureDocState port → a corrective (not supported)', async () => {
    const bridge = new FlexBridge('word'); // no captureDocState
    const { fetch } = scriptedFetch(['```cmd\noutline\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('show outline'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('outline not supported'),
    });
  });

  it('outline when captureDocState returns undefined → { outline: null }', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(undefined),
    });
    const { fetch } = scriptedFetch(['```cmd\noutline\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('show outline'));
    expect(readResults(events)[0]?.result).toEqual({ outline: null });
  });

  it('whole-document read (empty selector) falls back to captureDocState', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    // `read` with no selector → whole-document read intent.
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    const r = readResults(events)[0];
    expect(r?.intentLabel).toBe('read');
    expect(String(r?.result)).toContain('Heading One');
  });

  it('whole-document read returns { document: null } when capture yields undefined', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(undefined),
    });
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    expect(readResults(events)[0]?.result).toEqual({ document: null });
  });

  it('whole-document read with no captureDocState port → corrective', async () => {
    const bridge = new FlexBridge('word'); // no captureDocState
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('whole-document read not supported'),
    });
  });

  it('addressable read with a selector but no readRange port → corrective', async () => {
    const bridge = new FlexBridge('excel'); // no readRange
    const { fetch } = scriptedFetch(['```cmd\nread Sales!A1:B2\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read range'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('addressable read not supported'),
    });
  });

  it('addressable read with a readRange port → flattened slices', async () => {
    const calls: string[] = [];
    const bridge = new FlexBridge('excel', {
      readRange: (a1) => {
        calls.push(a1);
        return Promise.resolve(gfmTable());
      },
    });
    const { fetch } = scriptedFetch(['```cmd\nread Sales!A1:B2\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read range'));
    expect(calls).toEqual(['Sales!A1:B2']);
    expect(readResults(events)[0]?.result).toEqual([{ title: 'A1', text: expect.any(String) }]);
  });

  it('search with no searchDocument port → corrective', async () => {
    const bridge = new FlexBridge('word'); // no searchDocument
    const { fetch } = scriptedFetch(['```cmd\nsearch SLA floor\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('find'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('search not supported'),
    });
  });

  it('search dispatches to searchDocument and flattens the hits', async () => {
    const seen: string[] = [];
    const bridge = new FlexBridge('word', {
      searchDocument: (q) => {
        seen.push(q);
        return Promise.resolve([
          {
            ref: { id: 'hit:0', kind: 'paragraph', surface: 'word', title: 'Hit', live: true },
            value: { as: 'text', text: 'a found slice', mimeType: 'text/markdown' },
          },
        ]);
      },
    });
    const { fetch } = scriptedFetch(['```cmd\nsearch SLA floor\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('find'));
    expect(seen).toEqual(['SLA floor']);
    expect(readResults(events)[0]?.intentLabel).toBe('search SLA floor');
    expect(readResults(events)[0]?.result).toEqual([{ title: 'Hit', text: 'a found slice' }]);
  });

  it('a thrown read port degrades to a corrective { error }, never a thrown loop', async () => {
    const bridge = new FlexBridge('excel', {
      readRange: () => Promise.reject(new Error('host exploded')),
    });
    const { fetch } = scriptedFetch(['```cmd\nread Sales!A1:B2\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read range'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('read failed: host exploded'),
    });
  });
});

/* ─────────────────── effect-arg expressions for surface verbs ────────────── */

describe('AssistSession.runCommands — surface effect-arg expressions resolve at dry-run', () => {
  // A bridge that advertises the ADR-0006 effect kinds and serves a table read for composition.
  const composeBridge = () =>
    new FlexBridge('teams', {
      actuations: [
        { kind: 'post-message', surface: 'teams', title: 'Post', reversible: true },
        { kind: 'reply-mail', surface: 'teams', title: 'Mail', reversible: true },
        { kind: 'create-mail', surface: 'teams', title: 'Compose', reversible: true },
        { kind: 'append-page', surface: 'teams', title: 'Page', reversible: true },
      ],
      readRange: () => Promise.resolve(gfmTable()),
    });

  it('post (textExpr) resolves a composed scalar to concrete text', async () => {
    const bridge = composeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read X\n```',
      '```cmd\npost ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(session.runCommands('compose post', { approvePlan: () => true }));
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]!.kind).toBe('post-message');
    // 100 + 250 summed → the concrete resolved scalar lands as the post text.
    expect(bridge.applied[0]!.params.text).toBe('350');
  });

  it('mail (bodyExpr) resolves an expression at dry-run', async () => {
    const bridge = composeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read X\n```',
      '```cmd\nmail ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(session.runCommands('compose mail', { approvePlan: () => true }));
    expect(bridge.applied.map((a) => a.kind)).toEqual(['reply-mail']);
    const mail = bridge.applied[0]!.params.mail as { body?: string } | undefined;
    expect(mail?.body).toBe('350');
  });

  it('compose (bodyExpr) resolves an expression and keeps the literal subject', async () => {
    const bridge = composeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read X\n```',
      '```cmd\ncompose "Q3 follow-up" ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(session.runCommands('compose draft', { approvePlan: () => true }));
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]!.kind).toBe('create-mail');
    const mail = bridge.applied[0]!.params.mail as { subject?: string } | undefined;
    expect(mail?.subject).toBe('Q3 follow-up');
  });

  it('page (bodyExpr) resolves an expression and keeps the literal title', async () => {
    const bridge = composeBridge();
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read X\n```',
      '```cmd\npage "Notes" ($t | sum amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(session.runCommands('compose page', { approvePlan: () => true }));
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]!.kind).toBe('append-page');
    // The literal page title is carried as the target's matchText…
    const target = bridge.applied[0]!.params.target as { matchText?: string } | undefined;
    expect(target?.matchText).toBe('Notes');
    // …and the composed expression (100 + 250) resolves into the concrete page body.
    expect(bridge.applied[0]!.params.text).toBe('350');
  });
});

/* ─────────────────── plan approval — defensive paths ──────────────────────── */

describe('AssistSession.runCommands — approvePlan defensive + actuate failure', () => {
  const writingBridge = () =>
    new FlexBridge('excel', {
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
    });

  it('a THROWING approvePlan fails closed — the whole plan is blocked, nothing actuates', async () => {
    const bridge = writingBridge();
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const events = await collectLoop(
      session.runCommands('write', {
        approvePlan: () => {
          throw new Error('approver crashed');
        },
      }),
    );
    expect(bridge.applied).toHaveLength(0);
    expect(writeResults(events)[0]?.result.error?.code).toBe('plan_unapproved');
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('an async-rejecting approvePlan fails closed too', async () => {
    const bridge = writingBridge();
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const events = await collectLoop(
      session.runCommands('write', { approvePlan: () => Promise.reject(new Error('nope')) }),
    );
    expect(bridge.applied).toHaveLength(0);
    expect(writeResults(events)[0]?.result.error?.code).toBe('plan_unapproved');
    warn.mockRestore();
  });

  it('a thrown actuate becomes a corrective actuate_failed result (no thrown loop)', async () => {
    const bridge = writingBridge();
    bridge.actuateError = new Error('bridge boom');
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('write', { approvePlan: () => true }));
    const w = writeResults(events)[0];
    expect(w?.result.ok).toBe(false);
    expect(w?.result.error).toMatchObject({ code: 'actuate_failed', message: 'bridge boom' });
    expect(bridge.applied).toHaveLength(0);
  });
});

/* ─────────────────── dependency DAG enforcement (ADR-0008 §7) ─────────────── */

describe('AssistSession.runCommands — dependency DAG enforcement (ADR-0008 §7)', () => {
  // A bridge that FAILS write-cells (ok:false) but succeeds format-cells, recording every actuate
  // call so we can prove a skipped effect never reached the bridge.
  class DagBridge implements DocBridge {
    readonly surface = 'excel' as const;
    actuated: ActuationRequest[] = [];
    getCapabilities(): CapabilityManifest {
      return {
        surface: 'excel',
        contextKinds: [],
        actuations: [
          { kind: 'write-cells', surface: 'excel', title: 'set', reversible: true },
          { kind: 'format-cells', surface: 'excel', title: 'format', reversible: true },
        ],
      };
    }
    listContext(): Promise<ContextRef[]> {
      return Promise.resolve([]);
    }
    resolveContext(): Promise<ResolvedContext[]> {
      return Promise.resolve([]);
    }
    actuate(req: ActuationRequest): Promise<ActuationResult> {
      this.actuated.push(req);
      if (req.kind === 'write-cells') {
        return Promise.resolve({
          ok: false,
          changeId: req.changeId,
          kind: req.kind,
          error: { code: 'boom', message: 'cells failed' },
        });
      }
      return Promise.resolve({ ok: true, changeId: req.changeId, kind: req.kind });
    }
  }

  it('skips a dependent effect when its prerequisite fails; an independent effect still runs', async () => {
    const bridge = new DagBridge();
    // e1 set Report!A1 (write-cells) — will FAIL.
    // e2 format Report!A1:B2 — overlaps A1 ⇒ depends on e1 ⇒ must be SKIPPED, never actuated.
    // e3 format Other!Z9 — independent ⇒ must still run.
    const program =
      '```cmd\nset Report!A1 5\nformat Report!A1:B2 bold=true\nformat Other!Z9 bold=true\n```';
    const { fetch } = scriptedFetch([program, '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const results = writeResults(
      await collectLoop(session.runCommands('go', { approvePlan: () => true })),
    ).map((e) => e.result);

    expect(results[0]?.ok).toBe(false); // e1 actuated, failed
    expect(results[0]?.error?.code).toBe('boom');
    expect(results[1]?.ok).toBe(false); // e2 skipped (depends on the failed e1)
    expect(results[1]?.error?.code).toBe('prerequisite_failed');
    expect(results[2]?.ok).toBe(true); // e3 independent, ran

    // The skipped dependent NEVER reached the bridge; the failed prereq and the independent did.
    const ranges = bridge.actuated.map((r) => r.params.target?.range);
    expect(ranges).toContain('Report!A1');
    expect(ranges).toContain('Other!Z9');
    expect(ranges).not.toContain('Report!A1:B2');
  });

  it('plan-preview surfaces the distinct approval classes so authorities are not silently bundled (§H)', async () => {
    const bridge = new FlexBridge('excel', {
      actuations: [
        { kind: 'write-cells', surface: 'excel', title: 'set', reversible: true },
        { kind: 'post-message', surface: 'excel', title: 'post', reversible: true },
      ],
    });
    // An in-document edit + an external post in one plan — the preview must expose BOTH authorities.
    const { fetch } = scriptedFetch(['```cmd\nset A1 5\npost "hi"\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('go', { approvePlan: () => true }));
    const preview = events.find((e) => e.type === 'plan-preview') as Extract<
      CommandLoopEvent,
      { type: 'plan-preview' }
    >;
    expect(preview).toBeDefined();
    expect(preview.approvalClasses).toEqual(['in-document', 'external']);
    expect(preview.effects.map((e) => e.approvalClass)).toEqual(['in-document', 'external']);
  });
});

/* ─────────────────── command-block cap (truncation) ──────────────────────── */

describe('AssistSession.runCommands — per-turn command cap truncation', () => {
  it('truncates a block with more entries than maxCommandsPerTurn and yields a capped event', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    // Four reads, cap at 2 → block truncated, a `capped` event, and a corrective in results.
    const { fetch } = scriptedFetch([
      '```cmd\noutline\noutline\noutline\noutline\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('many', { maxCommandsPerTurn: 2 }));
    const capped = events.find(
      (e) => e.type === 'capped' && e.reason.includes('command block truncated'),
    );
    expect(capped).toBeDefined();
    // Only the first 2 of the 4 reads actually ran.
    expect(readResults(events).length).toBeLessThanOrEqual(2);
  });
});

/* ─────────────────── loop exhaustion + double no-fence ───────────────────── */

describe('AssistSession.runCommands — loop bounds', () => {
  it('emits `exhausted` when maxTurns is reached without `done`', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    // Always emit a read, never `done` → the loop hits the turn bound.
    const { fetch } = scriptedFetch(['```cmd\noutline\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('forever', { maxTurns: 2 }));
    const exhausted = events.find((e) => e.type === 'exhausted') as
      | Extract<CommandLoopEvent, { type: 'exhausted' }>
      | undefined;
    expect(exhausted?.turns).toBe(2);
  });

  it('two consecutive no-fence replies end the loop', async () => {
    const bridge = new FlexBridge('word');
    // No ```cmd block on either turn → re-prompt once, then break.
    const { fetch } = scriptedFetch(['just prose, no fence', 'still no fence']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('chat', { maxTurns: 8 }));
    const noFence = events.filter((e) => e.type === 'no-fence');
    // Re-prompted once after the first no-fence, then broke on the second — exactly two.
    expect(noFence.length).toBe(2);
    // It broke after the second no-fence rather than running all 8 turns (only 2 turns started).
    const turnStarts = events.filter((e) => e.type === 'turn-start');
    expect(turnStarts.length).toBe(2);
  });
});

/* ─────────────────── doc_state on command turns (nextCommandTurn) ────────── */

describe('AssistSession.runCommands — ambient doc_state feeds each command turn', () => {
  it('turn 2 carries a fresh <doc_state> alongside the result block', async () => {
    let captures = 0;
    const bridge = new FlexBridge('word', {
      captureDocState: () => {
        captures++;
        return Promise.resolve(snapshot());
      },
    });
    const { fetch, bodies } = scriptedFetch(['```cmd\noutline\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    // docState ON (default) so renderAmbientDocState runs for the command turns.
    const session = new AssistSession(bridge, client, { unit });

    await collectLoop(session.runCommands('go'));
    const turn2 = (bodies[1]!.query as { text?: string }).text ?? '';
    expect(turn2).toContain('```result');
    expect(turn2).toContain('<doc_state');
    expect(turn2).toContain('Continue. Next command?');
    expect(captures).toBeGreaterThanOrEqual(2);
  });

  it('a throwing captureDocState on a command turn is skipped, the loop still proceeds', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.reject(new Error('capture boom')),
    });
    const { fetch, bodies } = scriptedFetch([
      '```cmd\nsearch x\n```', // search not supported → corrective, but turn still completes
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const session = new AssistSession(bridge, client, { unit }); // docState ON

    const events = await collectLoop(session.runCommands('go'));
    // The turn-2 query has the result block but NO <doc_state> (capture threw).
    const turn2 = (bodies[1]!.query as { text?: string }).text ?? '';
    expect(turn2).toContain('```result');
    expect(turn2).not.toContain('<doc_state');
    expect(events.some((e) => e.type === 'done')).toBe(true);
    warn.mockRestore();
  });
});

/* ─────────────────── attachRef — public attach-by-chip path ──────────────── */

describe('AssistSession.attachRef', () => {
  it('resolves a ref through the bridge and adds every resolved part to the context', async () => {
    class AttachBridge extends FlexBridge {
      override resolveContext(): Promise<ResolvedContext[]> {
        return Promise.resolve([
          {
            ref: { id: 'sel:1', kind: 'selection', surface: 'word', title: 'Sel', live: true },
            value: { as: 'text', text: 'attached body', mimeType: 'text/markdown' },
          },
        ]);
      }
    }
    const bridge = new AttachBridge('word');
    const client = new StreamAssistClient(
      tokens,
      cfg,
      (async () => new Response('[]', { status: 200 })) as never,
    );
    const session = new AssistSession(bridge, client, { unit });
    expect(session.context.size).toBe(0);
    await session.attachRef({
      id: 'sel:1',
      kind: 'selection',
      surface: 'word',
      title: 'Sel',
      live: true,
    });
    expect(session.context.size).toBe(1);
  });
});

/* ─────────────────── help verb + composed slide/format in the loop ────────── */

describe('AssistSession.runCommands — control + render branches', () => {
  it('the `help` verb yields the grammar advertisement back as a result', async () => {
    const bridge = new FlexBridge('excel', {
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
    });
    const { fetch, bodies } = scriptedFetch(['```cmd\nhelp\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    await collectLoop(session.runCommands('what can you do'));
    // Turn 2's fed-back result block carries the grammar help text.
    const turn2 = (bodies[1]!.query as { text?: string }).text ?? '';
    expect(turn2).toContain('help');
    expect(turn2).toContain('done');
  });

  it('a literal-bullet slide effect previews the verbatim command and actuates the bullets', async () => {
    const bridge = new FlexBridge('powerpoint', {
      actuations: [
        { kind: 'insert-slide', surface: 'powerpoint', title: 'Slide', reversible: true },
      ],
    });
    const { fetch } = scriptedFetch([
      '```cmd\nslide "Roadmap" "ship v1" "ship v2"\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let preview: string | undefined;
    await collectLoop(
      session.runCommands('add a slide', {
        approvePlan: (effects) => {
          preview = effects[0]?.command;
          return true;
        },
      }),
    );
    // The plan preview renders the slide command line verbatim (title + quoted bullets).
    expect(preview).toBe('slide "Roadmap" "ship v1" "ship v2"');
    expect(bridge.applied).toHaveLength(1);
    const slide = bridge.applied[0]!.params.slide as { title?: string; bullets?: string[] };
    expect(slide.title).toBe('Roadmap');
    expect(slide.bullets).toEqual(['ship v1', 'ship v2']);
  });

  it('a slide whose bullets come from a large table is capped with a "+N more rows" tail', async () => {
    // 12 data rows → SLIDE_BULLET_CAP (10) bullets + one "+2 more rows" tail bullet.
    const manyRows = Array.from({ length: 12 }, (_, i) => `| R${i} | ${i} |`).join('\n');
    const gfm = `| region | amount |\n| --- | --- |\n${manyRows}`;
    const bridge = new FlexBridge('powerpoint', {
      actuations: [
        { kind: 'insert-slide', surface: 'powerpoint', title: 'Slide', reversible: true },
      ],
      readRange: () =>
        Promise.resolve([
          {
            ref: { id: 'r', kind: 'range', surface: 'excel', title: 'T', live: false },
            value: { as: 'text', text: gfm, mimeType: 'text/markdown' },
          },
        ]),
    });
    const { fetch } = scriptedFetch([
      '```cmd\nlet $t = read X\n```',
      '```cmd\nslide "All rows" ($t | select region,amount)\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    await collectLoop(session.runCommands('big slide', { approvePlan: () => true }));
    expect(bridge.applied).toHaveLength(1);
    const slide = bridge.applied[0]!.params.slide as { bullets?: string[] };
    expect(slide.bullets).toHaveLength(11); // 10 capped rows + the tail
    expect(slide.bullets?.[10]).toBe('(+2 more rows)');
  });

  it('a format effect renders its key=value props in the plan preview', async () => {
    const bridge = new FlexBridge('excel', {
      actuations: [{ kind: 'format-cells', surface: 'excel', title: 'Format', reversible: true }],
    });
    const { fetch } = scriptedFetch([
      '```cmd\nformat Sales!A1:B1 bold=true fill=#FFF2CC\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    let preview: string | undefined;
    await collectLoop(
      session.runCommands('format header', {
        approvePlan: (effects) => {
          preview = effects[0]?.command;
          return true;
        },
      }),
    );
    expect(preview).toBe('format Sales!A1:B1 bold=true fill=#FFF2CC');
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]!.kind).toBe('format-cells');
  });
});

/* ─────────────────── apply() — provenance stamping branch ────────────────── */

describe('AssistSession.apply — provenance presence/absence', () => {
  it('omits provenance entirely when no turn has produced any (no stamp)', async () => {
    const bridge = new FlexBridge('word', {
      actuations: [{ kind: 'tracked-change', surface: 'word', title: 't', reversible: true }],
    });
    const client = new StreamAssistClient(
      tokens,
      cfg,
      (async () => new Response('[]', { status: 200 })) as never,
    );
    const session = new AssistSession(bridge, client, { unit });
    const result = await session.apply('tracked-change', { text: 'x' }, asChangeId('c1'));
    expect(result.ok).toBe(true);
    expect(bridge.applied[0]!.provenance).toBeUndefined();
  });
});
