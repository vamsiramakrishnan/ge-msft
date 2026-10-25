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
import { asChangeId, asSessionId } from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import type { DocStateSnapshot } from '@ge/contracts';
import { AssistSession } from './assist-session.js';
import type { CommandLoopEvent } from './assist-session.js';
import type { DocBridge } from './bridge.js';
import { RuntimeHooks } from './hooks.js';

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
  readDocumentPage?: DocBridge['readDocumentPage'];
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

  it('whole-document read returns body text and a continuation instead of the outline', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    bridge.readDocumentPage = async () => ({
      scope: 'body',
      units: [{ id: 'p1', text: 'Actual paragraph body' }],
      nextCursor: '40',
      limitations: [],
    });
    // `read` with no selector → whole-document read intent.
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    const r = readResults(events)[0];
    expect(r?.intentLabel).toBe('read document');
    expect(r?.result).toMatchObject({
      units: [{ id: 'p1', text: 'Actual paragraph body' }],
      complete: false,
      next: 'read document:40',
    });
  });

  it('whole-document read does not pass an absent outline off as document content', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(undefined),
    });
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('Paged document reads are unavailable'),
    });
  });

  it('whole-document read with no captureDocState port → corrective', async () => {
    const bridge = new FlexBridge('word'); // no captureDocState
    const { fetch } = scriptedFetch(['```cmd\nread\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('read all'));
    expect(readResults(events)[0]?.result).toMatchObject({
      error: expect.stringContaining('Paged document reads are unavailable'),
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

  it('counts a "3x3" table as 3 rows including the header before approval (user decision 2026-10-17)', async () => {
    const bridge = new FlexBridge('powerpoint', {
      actuations: [
        { kind: 'add-table-slide', surface: 'powerpoint', title: 'Table', reversible: true },
      ],
    });
    const long =
      '/add-table-slide slide=4 rows="No\\tDimensi\\tCatatan\\n1\\tA\\ta\\n2\\tB\\tb\\n3\\tC\\tc"';
    const right = '/add-table-slide slide=4 rows="No\\tDimensi\\tCatatan\\n1\\tA\\ta\\n2\\tB\\tb"';
    const { fetch, bodies } = scriptedFetch([
      '```cmd\n' + long + '\n```',
      '```cmd\n' + right + '\n```',
      '```cmd\ndone\n```',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(
      session.runCommands('pada slide no 4 ini, generate table 3x3', { approvePlan: () => true }),
    );
    expect(JSON.stringify(bodies[1])).toContain('3 rows in total, the header row included');
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]!.params.tableGrid!.rows).toHaveLength(3);
  });

  it('adds the header on top only when the user asks for data rows (3x3 with 3 data rows)', async () => {
    const bridge = new FlexBridge('powerpoint', {
      actuations: [
        { kind: 'add-table-slide', surface: 'powerpoint', title: 'Table', reversible: true },
      ],
    });
    const four =
      '/add-table-slide slide=4 rows="No\\tDimensi\\tCatatan\\n1\\tA\\ta\\n2\\tB\\tb\\n3\\tC\\tc"';
    const { fetch } = scriptedFetch(['```cmd\n' + four + '\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    await collectLoop(
      session.runCommands('generate table 3x3 with 3 data rows below the header', {
        approvePlan: () => true,
      }),
    );
    expect(bridge.applied[0]!.params.tableGrid!.rows).toHaveLength(4);
  });

  it('keeps writes that leave the document (Teams posts) at 8 per turn under the 30 cap', async () => {
    const bridge = composeBridge();
    const posts = Array.from({ length: 9 }, (_, i) => `post "update ${i}"`).join('\n');
    const { fetch } = scriptedFetch(['```cmd\n' + posts + '\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    const events = await collectLoop(
      session.runCommands('post nine updates', { approvePlan: () => true }),
    );
    expect(bridge.applied).toHaveLength(8);
    expect(events.find((e) => e.type === 'capped')).toMatchObject({ reason: 'write cap 8/turn' });
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

  it('a thrown host dispatch becomes an uncertain corrective result (no thrown loop)', async () => {
    const bridge = writingBridge();
    bridge.actuateError = new Error('bridge boom');
    const { fetch } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('write', { approvePlan: () => true }));
    const w = writeResults(events)[0];
    expect(w?.result.ok).toBe(false);
    expect(w?.result.error?.code).toBe('outcome_unknown');
    expect(w?.result.recoveryPending).toBe(true);
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

  it('a no-fence event carries a bounded, redacted snippet of the unparsed reply', async () => {
    const bridge = new FlexBridge('word');
    const { fetch } = scriptedFetch(['just prose, no fence', 'still no fence']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('chat', { maxTurns: 8 }));
    const noFence = events.find((e) => e.type === 'no-fence') as
      | Extract<CommandLoopEvent, { type: 'no-fence' }>
      | undefined;
    expect(noFence?.rawSnippet).toBeDefined();
    expect(noFence!.rawSnippet!.length).toBeLessThanOrEqual(200);
  });

  it('a first no-fence landing on the LAST budgeted turn still gets its one recovery turn', async () => {
    const bridge = new FlexBridge('word', {
      captureDocState: () => Promise.resolve(snapshot()),
    });
    // Turn 1 is a normal read (not a no-fence, not `done`). maxTurns is 2, so turn 2 is the last
    // budgeted turn — the FIRST no-fence lands exactly there. Turn 3 (the one-turn recovery
    // reprompt) is also no-fence, so the loop breaks on the second no-fence rather than running
    // forever.
    const { fetch } = scriptedFetch([
      '```cmd\noutline\n```',
      'just prose, no fence',
      'still no fence too',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('chat', { maxTurns: 2 }));
    const noFence = events.filter((e) => e.type === 'no-fence');
    // Proves the recovery turn actually ran: TWO no-fence events, not one.
    expect(noFence.length).toBe(2);
    // Proves the loop got exactly one extra turn past maxTurns to use the recovery reprompt.
    const turnStarts = events.filter((e) => e.type === 'turn-start');
    expect(turnStarts.length).toBe(3); // maxTurns (2) + 1 recovery turn
  });

  it('`exhausted.turns` reflects the actual last-started turn, not a hardcoded maxTurns', async () => {
    const bridge = new FlexBridge('word');
    // A program ran at turn 1; then two consecutive no-fence replies end the loop at turn 3, well
    // before maxTurns (6).
    const { fetch } = scriptedFetch([
      '```cmd\noutline\n```',
      'just prose, no fence',
      'still no fence',
    ]);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('chat', { maxTurns: 6 }));
    const exhausted = events.find((e) => e.type === 'exhausted') as
      | Extract<CommandLoopEvent, { type: 'exhausted' }>
      | undefined;
    expect(exhausted?.turns).toBe(3);
  });

  it('keeps a confirmed change plan incomplete when the model only replies in prose', async () => {
    const bridge = new FlexBridge('word');
    const { fetch } = scriptedFetch(['I will rewrite it.', 'Rewriting the paragraph now.']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });
    const task =
      '<confirmed_plan>\noriginal_request: rewrite the intro\nintent: rewrite\nstep 1: rewrite the introduction paragraph\n</confirmed_plan>';

    const events = await collectLoop(session.runCommands(task, { maxTurns: 6 }));
    expect(events.some((e) => e.type === 'done' && 'answerOnly' in e)).toBe(false);
    expect(events.some((e) => e.type === 'exhausted')).toBe(true);
  });

  it('ends as an answer when the model replies in prose twice and no command ran', async () => {
    const bridge = new FlexBridge('word');
    const { fetch } = scriptedFetch(['Hello! How can I help?', 'Hi! What would you like to do?']);
    const client = new StreamAssistClient(tokens, cfg, fetch);
    const session = new AssistSession(bridge, client, { unit, context: { docState: false } });

    const events = await collectLoop(session.runCommands('hii', { maxTurns: 6 }));
    expect(events.some((e) => e.type === 'exhausted')).toBe(false);
    expect(events.at(-1)).toMatchObject({
      type: 'done',
      answerOnly: true,
      answer: 'Hi! What would you like to do?',
    });
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
    // An unreadable document is said, never silently dropped (the model invented figures otherwise).
    expect(turn2).toContain('<doc_state unavailable="true">');
    expect(turn2).not.toContain('<doc_state surface=');
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

describe('complete document summary integration', () => {
  it('reads every page twice, isolates the wire, and preserves the existing chat session', async () => {
    const bridge = new FlexBridge('word');
    const reads: Array<string | undefined> = [];
    bridge.readDocumentPage = async ({ cursor }) => {
      reads.push(cursor);
      return {
        scope: 'Word main-body paragraphs',
        limitations: ['Images excluded.'],
        units: [
          {
            id: cursor ?? 'first',
            text: cursor ? 'Final paragraph evidence' : 'First paragraph evidence',
          },
        ],
        ...(cursor ? {} : { nextCursor: '40' }),
      };
    };
    const { fetch, bodies } = scriptedFetch(['Summary of both paragraphs.']);
    const client = new StreamAssistClient(
      tokens,
      {
        ...cfg,
        skills: ['projects/p/agents/commander'],
        dataStores: ['external-store'],
      },
      fetch,
    );
    const session = new AssistSession(bridge, client, {
      unit: { ...unit, surfaceContext: { kind: 'word', selection: 'STALE SELECTION' } },
      resumeSessionId: asSessionId('existing-chat'),
    });
    session.context.add({
      ref: { id: 'old-attachment', kind: 'document', surface: 'word', title: 'Old attachment' },
      value: { as: 'text', text: 'STALE ATTACHMENT' },
    });
    const events = await collectLoop(session.ask('read the entire doc and summarise it'));
    expect(reads).toEqual([undefined, '40', undefined, '40']);
    expect(bodies).toHaveLength(1);
    expect(bodies[0]).toMatchObject({ isSessionLess: true });
    expect(bodies[0]).not.toHaveProperty('session');
    expect(bodies[0]).not.toHaveProperty('agentsSpec');
    expect(bodies[0]).not.toHaveProperty('toolsSpec');
    expect(JSON.stringify(bodies)).toContain('Final paragraph evidence');
    expect(JSON.stringify(bodies)).not.toContain('STALE');
    expect(events.filter((event) => event.type === 'token')).toHaveLength(1);
    expect(events.some((event) => event.type === 'provenance')).toBe(false);
    expect(session.sessionId).toBe('existing-chat');
    expect(session.executions.list()[0]?.status).toBe('completed');
    expect(bridge.applied).toEqual([]);
  });

  it('honors a read guard before sending source content and records a blocked task', async () => {
    const bridge = new FlexBridge('word');
    bridge.readDocumentPage = vi.fn();
    const { fetch, bodies } = scriptedFetch(['must not be called']);
    const hooks = new RuntimeHooks();
    hooks.register({
      id: 'stop-read',
      on: 'tool:before',
      mode: 'guard',
      handle: () => ({ kind: 'block', reason: 'Read denied.' }),
    });
    const session = new AssistSession(bridge, new StreamAssistClient(tokens, cfg, fetch), {
      unit,
      hooks,
    });
    await expect(collectLoop(session.ask('summarize this document'))).rejects.toThrow(
      'Read denied',
    );
    expect(bridge.readDocumentPage).not.toHaveBeenCalled();
    expect(bodies).toEqual([]);
    expect(session.executions.list()[0]?.status).toBe('blocked');
  });

  it('rejects conflicting uploaded grounding without silently ignoring it', async () => {
    const bridge = new FlexBridge('word');
    bridge.readDocumentPage = vi.fn();
    const { fetch, bodies } = scriptedFetch(['must not be called']);
    const session = new AssistSession(bridge, new StreamAssistClient(tokens, cfg, fetch), { unit });
    const events = await collectLoop(
      session.ask('summarize this document', { grounding: { fileIds: ['uploaded-file'] } }),
    );
    expect(events).toContainEqual(
      expect.objectContaining({ type: 'error', code: 'document_scope_conflict' }),
    );
    expect(bridge.readDocumentPage).not.toHaveBeenCalled();
    expect(bodies).toEqual([]);
    expect(session.executions.list()[0]?.status).toBe('incomplete');
  });
});

describe('paged document composition', () => {
  it('refuses to bind a partial document as a complete pipeline source', async () => {
    const bridge = new FlexBridge('word');
    bridge.readDocumentPage = async () => ({
      scope: 'Body',
      units: [{ id: 'first', text: 'Only the first page' }],
      nextCursor: '40',
      limitations: [],
    });
    const { fetch } = scriptedFetch(['```cmd\nlet $t = read document\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, new StreamAssistClient(tokens, cfg, fetch), { unit });
    const events = await collectLoop(session.runCommands('read with a pipeline'));
    expect(JSON.stringify(events)).toContain('Paged document read is incomplete');
    expect(bridge.applied).toEqual([]);
  });
});

describe('AssistSession evidence-bound execution', () => {
  function setup() {
    const bridge = new FlexBridge('excel', {
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
    });
    let definition = 'Revenue excludes refunds';
    bridge.readDocumentPage = async () => ({
      scope: 'workbook',
      limitations: [],
      units: [
        { id: 'A1', text: '100', relatedIds: ['definition'] },
        { id: 'definition', text: definition },
      ],
    });
    const { fetch, bodies } = scriptedFetch(['```cmd\nset A1 1\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, new StreamAssistClient(tokens, cfg, fetch), {
      unit,
      context: { docState: false },
    });
    return {
      bridge,
      session,
      bodies,
      change: () => {
        definition = 'Revenue includes refunds';
      },
    };
  }
  it('rechecks non-target evidence after approval, before host dispatch', async () => {
    const { session, bridge, change } = setup();
    const binding = await session.prepareEvidence(
      { requiredIds: ['A1'] },
      { authorizeEffect: () => true },
    );
    // Completion can fail freshness as well; neither path may dispatch the write.
    try {
      await collectLoop(
        session.runCommands('write', {
          evidenceBinding: binding.id,
          approvePlan: () => {
            change();
            return true;
          },
        }),
      );
    } catch (error) {
      expect(String(error)).toContain('changed');
    }
    expect(bridge.applied).toHaveLength(0);
  });
  it('defaults to no write authority even when the user approves a plan', async () => {
    const { session, bridge } = setup();
    const binding = await session.prepareEvidence({ requiredIds: ['A1'] });
    await collectLoop(
      session.runCommands('write', { evidenceBinding: binding.id, approvePlan: () => true }),
    );
    expect(bridge.applied).toHaveLength(0);
  });
  it('dispatches an authorized effect once and consumes its freshness binding', async () => {
    const { session, bridge, bodies } = setup();
    const binding = await session.prepareEvidence(
      { requiredIds: ['A1'] },
      { authorizeEffect: (request) => request.kind === 'write-cells' },
    );
    await collectLoop(
      session.runCommands('write', { evidenceBinding: binding.id, approvePlan: () => true }),
    );
    expect(bridge.applied).toHaveLength(1);
    expect(JSON.stringify(bodies[0])).toContain('Revenue excludes refunds');
    await expect(
      collectLoop(
        session.runCommands('again', { evidenceBinding: binding.id, approvePlan: () => true }),
      ),
    ).rejects.toThrow('consumed');
  });
  it('rejects released bindings before invoking the model', async () => {
    const { session, bodies } = setup();
    const binding = await session.prepareEvidence({ requiredIds: ['A1'] });
    session.releaseEvidenceBinding(binding.id);
    await expect(
      collectLoop(session.runCommands('write', { evidenceBinding: binding.id })),
    ).rejects.toThrow('released');
    expect(bodies).toHaveLength(0);
  });
});

describe('AssistSession absolute DocFs reads advertised by the skills', () => {
  it('routes evidence grep and Unicode text continuations through the command loop', async () => {
    const bridge = new FlexBridge('word');
    const { fetch } = scriptedFetch(['```cmd\ndone\n```']);
    const session = new AssistSession(bridge, new StreamAssistClient(tokens, cfg, fetch), {
      unit,
      context: { docState: false },
    });
    const text = 'Governed revenue definition: ' + '😀'.repeat(2500);
    const { id } = await session.sourceWorkspace.capture({
      id: 'source',
      readPage: async () => ({
        scope: 'body',
        limitations: ['body text only'],
        units: [{ id: 'definition', text }],
      }),
    });
    const grepEvents = await collectLoop(
      session.runCommandProgram(`grep /evidence/${id}/units "Governed revenue"`),
    );
    const search = readResults(grepEvents)[0]!.result as { matches: Array<{ path: string }> };
    expect(search.matches[0]!.path).toBe(`/evidence/${id}/units/0/0.json`);
    let nextPath: string | undefined = `/evidence/${id}/windows/0/0.json`;
    let exact = '';
    while (nextPath) {
      const events = await collectLoop(session.runCommandProgram(`cat ${nextPath}`));
      const result = readResults(events)[0]!.result as { text: string; truncated: boolean };
      expect(result.truncated).toBe(false);
      const page = JSON.parse(result.text) as { text: string; nextPath?: string };
      exact += page.text;
      nextPath = page.nextPath;
    }
    expect(exact).toBe(text);
    expect(bridge.applied).toEqual([]);
  });
  it('preserves named workspace reads and reports missing paths and line truncation', async () => {
    const { fetch } = scriptedFetch(['```cmd\ndone\n```']);
    const session = new AssistSession(
      new FlexBridge('word'),
      new StreamAssistClient(tokens, cfg, fetch),
      { unit, context: { docState: false }, skillFiles: { 'lines.md': 'first\nsecond' } },
    );
    await collectLoop(session.runCommandProgram('save notes.md = "first\\nsecond"'));
    const legacy = readResults(await collectLoop(session.runCommandProgram('cat notes.md')))[0]!
      .result as { workspace: string; preview: string };
    expect(legacy.workspace).toBe('cat');
    expect(legacy.preview).toContain('second');
    const absolute = readResults(
      await collectLoop(session.runCommandProgram('cat /skills/lines.md head=1')),
    )[0]!.result as { text: string; truncated: boolean };
    expect(absolute).toMatchObject({ text: 'first', truncated: true });
    const missing = readResults(
      await collectLoop(session.runCommandProgram('cat /evidence/missing/units/0/0.json')),
    )[0]!.result as { error: string };
    expect(missing.error).toContain('no such file');
  });
});
