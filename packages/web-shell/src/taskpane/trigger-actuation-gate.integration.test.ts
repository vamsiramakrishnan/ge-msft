// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { asChangeId } from '@ge/contracts';
import type { ActuationParams, ActuationRequest, ProvenancePayload } from '@ge/contracts';
import { AssistSession } from '@ge/runtime';
import { ExcelBridge } from '@ge/bridge-excel';
import {
  CONTINUE,
  EventBus,
  TriggerRegistry,
  debounce,
  type HostEvent,
  type Scheduler,
} from '@ge/triggers';
import { installFakeExcel, excelSeed, scriptedClient } from '../test-harness/index.js';
import type { ExcelSimulator } from '../test-harness/index.js';

/**
 * INTERPLAY — trigger-actuation-gate.
 *
 * Wires the REAL packages across their seam and mocks ONLY the outer boundaries:
 *   - the Office HOST → the in-memory fake Excel host (`installFakeExcel`, sets `globalThis.Excel`).
 *   - the Google NETWORK → the scripted fake `StreamAssistClient` (`scriptedClient`).
 *
 * Everything BETWEEN is real:
 *   - real `ExcelBridge` (its `watch()` emits HostEvents; its `actuate('write-cells')` mutates the host),
 *   - real `EventBus` + real `debounce` (the event-source pipeline that coalesces rapid emissions),
 *   - real `TriggerRegistry` (the `pre-actuation` veto gate that runs inside `AssistSession.apply`),
 *   - real `AssistSession` (which stamps last-turn provenance and runs the trigger gate before the bridge actuates).
 *
 * We assert OBSERVABLE cross-boundary behavior, read back from the fake host:
 *   1. rapid HostEvents from the real bridge `watch()` are debounced/coalesced to ONE registry dispatch,
 *   2. a `pre-actuation` gate trigger HOLDS the effect — the host cell is NOT mutated while the gate is pending,
 *      and the cell IS written ONLY after the gate resolves to `continue`,
 *   3. a `block` outcome stops the write entirely (host untouched, error code `blocked`),
 *   4. the never-resolving gate path: `apply()` stays pending and the host is never mutated.
 */

/** A single-sheet workbook so a `write-cells` lands in a deterministic, readable cell. */
function gateSeed(): ReturnType<typeof excelSeed> {
  return excelSeed({
    sheets: [
      {
        name: 'Sales',
        origin: 'A1',
        values: [
          ['metric', 'value'],
          ['total revenue', ''],
        ],
      },
    ],
    activeSheet: 'Sales',
    selection: 'Sales!A1:B1',
  });
}

/** Read the cell at `Sales!B2` (row index 1, col index 1) back out of the fake host snapshot. */
function cellB2(sim: ExcelSimulator): string {
  return sim.snapshot().sheets.find((s) => s.name === 'Sales')?.values[1]?.[1] ?? '';
}

/** A deterministic, hand-driven scheduler for `debounce` (no real timers). */
function manualScheduler(): { scheduler: Scheduler; fire: () => void; pending: () => boolean } {
  let fn: (() => void) | null = null;
  return {
    scheduler: {
      set: (cb) => {
        fn = cb;
        return Symbol('handle');
      },
      clear: () => {
        fn = null;
      },
    },
    fire: () => {
      const f = fn;
      fn = null;
      f?.();
    },
    pending: () => fn !== null,
  };
}

/** The write the agent proposes: set `Sales!B2` to a computed total. A fresh copy per call. */
function writeParams(): ActuationParams {
  return { target: { range: 'Sales!B2' }, cells: [['1180']] };
}

/** A scripted turn that stamps real provenance so an approved write records who/what into the host. */
const SCRIPT = ['Computed the total.', '```cmd\ndone\n```'];

let sim: ExcelSimulator | undefined;
let unsubscribe: (() => void) | undefined;

beforeEach(() => {
  vi.useRealTimers();
});

afterEach(() => {
  unsubscribe?.();
  sim?.restore();
  unsubscribe = undefined;
  sim = undefined;
});

/** Build a live session over the real bridge + scripted client + the given trigger registry. */
function makeSession(registry: TriggerRegistry): { bridge: ExcelBridge; session: AssistSession } {
  const bridge = new ExcelBridge();
  const { client } = scriptedClient(SCRIPT);
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' } },
    triggers: registry,
  });
  return { bridge, session };
}

/** Pin this answer's attribution to its explicit proposal, independent of later session turns. */
async function proposalProvenance(session: AssistSession): Promise<ProvenancePayload> {
  let provenance: ProvenancePayload | undefined;
  for await (const event of session.ask('compute the total revenue')) {
    if (event.type === 'provenance') provenance = event.payload;
  }
  if (!provenance) throw new Error('The scripted answer must provide proposal attribution.');
  return provenance;
}

describe('INTERPLAY: triggers + runtime + fake host — debounce/coalesce', () => {
  it('coalesces rapid bridge watch() events into ONE registry dispatch', async () => {
    sim = installFakeExcel(gateSeed());
    const bridge = new ExcelBridge();
    const bus = new EventBus();
    const registry = new TriggerRegistry();
    const sched = manualScheduler();

    const dispatched: HostEvent[] = [];
    registry.register({
      id: 'ambient',
      on: 'selection-changed',
      handle: (e) => {
        dispatched.push(e);
        return CONTINUE;
      },
    });

    // Real event-source pipeline: bridge.watch() → bus → real debounce → real registry.dispatch.
    const onSelection = debounce(
      (e: HostEvent) => {
        bus.emit(e);
      },
      150,
      sched.scheduler,
    );
    bus.on('selection-changed', (e) => void registry.dispatch(e));

    // The REAL ExcelBridge.watch() registers the Office event handlers on the fake host.
    unsubscribe = bridge.watch((e) => onSelection(e));
    // Let the bridge's Excel.run() registration settle before firing.
    await Promise.resolve();
    await Promise.resolve();

    // Fire three rapid selection moves through the REAL bridge's wired Office handlers.
    sim.events.fireSelectionChanged('Sales!A1');
    sim.events.fireSelectionChanged('Sales!A2');
    sim.events.fireSelectionChanged('Sales!B3');

    // Nothing dispatched yet — the debounce window has not elapsed (coalescing in flight).
    expect(dispatched).toHaveLength(0);
    expect(sched.pending()).toBe(true);

    // The trailing edge fires once; only the LAST event survives.
    sched.fire();
    await Promise.resolve();

    expect(dispatched).toHaveLength(1);
    expect((dispatched[0] as { preview?: string }).preview).toBe('Sales!B3');
  });
});

describe('INTERPLAY: triggers + runtime + fake host — the actuation gate holds the effect', () => {
  it('HOLDS the host mutation while the gate is pending, then writes ONLY after it opens', async () => {
    sim = installFakeExcel(gateSeed());
    const registry = new TriggerRegistry();

    // A deferred pre-actuation gate: the trigger does not decide until the test "approves".
    let openGate!: () => void;
    const gateOpened = new Promise<void>((resolve) => {
      openGate = resolve;
    });
    let gateSawRequest: ActuationRequest | undefined;
    let sawGate!: () => void;
    const gateEntered = new Promise<void>((resolve) => {
      sawGate = resolve;
    });
    registry.register({
      id: 'deferred-approval',
      on: 'pre-actuation',
      handle: async (e) => {
        if (e.type === 'pre-actuation') gateSawRequest = e.request;
        sawGate();
        await gateOpened; // hold the effect until the test opens the gate
        return CONTINUE;
      },
    });

    const { session } = makeSession(registry);
    const provenance = await proposalProvenance(session);

    // Kick off the gated write but DO NOT await it yet — it suspends inside the trigger gate.
    let result: Awaited<ReturnType<AssistSession['apply']>> | undefined;
    const applying = session
      .apply('write-cells', writeParams(), asChangeId('gate-write-1'), provenance)
      .then((r) => (result = r));

    // Synchronize with gate entry, independently of the number of lifecycle-hook awaits.
    await gateEntered;

    // OBSERVABLE: the gate received the real request, but the host cell is STILL EMPTY — held.
    expect(gateSawRequest?.kind).toBe('write-cells');
    expect(gateSawRequest?.params.target?.range).toBe('Sales!B2');
    expect(cellB2(sim)).toBe('');
    expect(result).toBeUndefined();

    // Open the gate → the real bridge actuates → the host cell is finally written.
    openGate();
    await applying;

    expect(result?.ok).toBe(true);
    expect(cellB2(sim)).toBe('1180');
    // Provenance from the reviewed proposal landed durably in the host settings bag (not dropped/missing).
    expect(result?.provenanceMissing).toBeUndefined();
    // CROSS-BOUNDARY READBACK: read the durable provenance record straight out of the fake host's
    // workbook settings bag (where `ExcelBridge.persistProvenance` wrote it via
    // `Office.context.document.settings.set`). The record is keyed `ge:prov:<changeId>` and carries
    // the runtime-stamped identity — proof the gated write traversed runtime → bridge → host metadata.
    expect(sim.office.settingsSaved).toBe(true);
    const persisted = sim.office.settings.get('ge:prov:gate-write-1');
    expect(persisted).toBeTypeOf('string');
    const record = JSON.parse(persisted as string) as { changeId?: string; sources?: unknown };
    expect(record.changeId).toBe('gate-write-1');
  });

  it('a `block` gate outcome stops the write entirely — host untouched, error code "blocked"', async () => {
    sim = installFakeExcel(gateSeed());
    const registry = new TriggerRegistry();
    registry.register({
      id: 'veto',
      on: 'pre-actuation',
      handle: () => ({ kind: 'block', reason: 'residency policy: cell write withheld' }),
    });

    const { session } = makeSession(registry);
    const provenance = await proposalProvenance(session);

    const result = await session.apply(
      'write-cells',
      writeParams(),
      asChangeId('gate-write-blocked'),
      provenance,
    );

    // The gate vetoed BEFORE the bridge ran: the host cell is untouched and the error surfaces.
    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe('blocked');
    expect(result.error?.message).toContain('residency policy');
    expect(cellB2(sim)).toBe('');
  });

  it('NEVER-RESOLVES gate path: apply() stays pending and the host is never mutated', async () => {
    sim = installFakeExcel(gateSeed());
    const registry = new TriggerRegistry();
    // A gate that never decides (the "hung approval" / timeout path).
    registry.register({
      id: 'never-resolves',
      on: 'pre-actuation',
      handle: () => new Promise(() => undefined),
    });

    const { session } = makeSession(registry);
    const provenance = await proposalProvenance(session);

    let settled = false;
    const applying = session
      .apply('write-cells', writeParams(), asChangeId('gate-write-hung'), provenance)
      .then(() => (settled = true));

    // Race the never-resolving apply against a short timeout: the timeout must win.
    const timeout = new Promise<'timeout'>((resolve) => setTimeout(() => resolve('timeout'), 20));
    const winner = await Promise.race([applying.then(() => 'applied' as const), timeout]);

    expect(winner).toBe('timeout');
    expect(settled).toBe(false);
    // The host cell is NEVER written while the gate is hung — the effect is held indefinitely.
    expect(cellB2(sim)).toBe('');
  });
});

describe('INTERPLAY: triggers + runtime + fake host — post-actuation audit fires after the write', () => {
  it('an approved write fires a real post-actuation audit carrying the landed result', async () => {
    sim = installFakeExcel(gateSeed());
    const registry = new TriggerRegistry();

    const audited: HostEvent[] = [];
    registry.register({
      id: 'auditor',
      on: 'post-actuation',
      handle: (e) => {
        audited.push(e);
        return CONTINUE;
      },
    });

    const { session } = makeSession(registry);
    const provenance = await proposalProvenance(session);

    const result = await session.apply(
      'write-cells',
      writeParams(),
      asChangeId('gate-write-audit'),
      provenance,
    );

    // Let the fire-and-forget post-actuation dispatch run.
    await Promise.resolve();
    await Promise.resolve();

    expect(result.ok).toBe(true);
    expect(cellB2(sim)).toBe('1180');
    // The audit hook saw the post-actuation event AFTER the host mutation landed.
    expect(audited).toHaveLength(1);
    const event = audited[0];
    expect(event?.type).toBe('post-actuation');
    if (event?.type === 'post-actuation') {
      expect(event.request.changeId).toBe('gate-write-audit');
      expect(event.result.ok).toBe(true);
    }
  });
});
