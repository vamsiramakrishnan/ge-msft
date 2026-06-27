// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  installFakeExcel,
  excelSeed,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * INTERPLAY — grammar-effect-parity.
 *
 * Seam wired (all REAL, no stubs between the boundaries): the `@ge/contracts` command/expression
 * grammar (`parseProgramBlock` → `ParsedExpr`/`ParsedCommand`/`CommandParseError`) → the `@ge/runtime`
 * orchestrator's plan machinery (`AssistSession.runCommands`: pure pipelines evaluated inline +
 * effects type-checked, dry-run-resolved into a gated `PlanEffect[]`) → the REAL `@ge/bridge-excel`
 * typed effect (`actuate({ kind: 'write-cells' })` mutating the in-memory host). Only the OUTER
 * boundaries are faked: the Office host (in-memory `installFakeExcel`) and the model stream
 * (`scriptedClient`).
 *
 * The track's thesis, asserted across the seam:
 *   1. PURE transforms (pipelines / `let` bindings) compose WITHOUT gating — they never produce a
 *      plan slot, never stage a `pendingPlan`, never touch the bridge — yet their COMPUTED value
 *      flows into an effect that IS gated and counted (one write in the plan summary), and only the
 *      single approval lands the computed literal in the host.
 *   2. A malformed / unknown command is rejected by the GRAMMAR (a `CommandParseError` in the
 *      program block) before it can become an effect — so it never reaches the bridge — while the
 *      valid effect in the same turn still forms the (single-effect) plan and actuates.
 *   3. A structurally-invalid effect arg (a non-scalar `table` Value written into one cell) is
 *      rejected by the orchestrator's dry-run, not the bridge: it never enters the plan, never
 *      actuates, and the pure pipeline beside it still composes free.
 *
 * "Observable across the boundary" everywhere: we read the mutated fake host back (or assert it is
 * untouched), and read the staged plan/summary off the live controller state + rendered DOM.
 */

let sim: ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

/** A small, deterministic workbook: an East/West revenue grid + an empty Summary target cell. */
function paritySeed(): ReturnType<typeof excelSeed> {
  return excelSeed({
    sheets: [
      {
        name: 'Sales',
        origin: 'A1',
        values: [
          ['region', 'rep', 'revenue', 'cost'],
          ['East', 'Alice', '300', '120'],
          ['East', 'Bob', '250', '100'],
          ['West', 'Carol', '180', '90'],
        ],
      },
      {
        name: 'Summary',
        origin: 'A1',
        values: [
          ['metric', 'value'],
          ['east revenue', ''],
        ],
      },
    ],
    activeSheet: 'Sales',
    selection: 'Sales!A2:D2',
  });
}

describe('grammar ↔ orchestrator ↔ bridge effect parity', () => {
  it('pure pipelines compose FREE while the effect that consumes them is gated + counted once', async () => {
    sim = installFakeExcel(paritySeed());
    // Two PURE entries (a `let` binding + a terminal scalar pipeline) and ONE effect that consumes
    // the bound value. The pure transforms run inline over the SEEDED grid (East = 300 + 250 = 550)
    // with NO gate; only the `set` becomes a plan slot.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n' +
          'let $east = read Sales!A1:D4 | filter region=East\n' +
          '$east | sum revenue\n' +
          'set Summary!B2 = ($east | sum revenue)\n' +
          '```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total East revenue into the summary');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // PARITY: the plan holds exactly ONE effect — the two pure entries produced NO plan slot, so
    // they were never gated. The summary counts a single write.
    const plan = ui!.controller.getState().pendingPlan!;
    expect(plan.effects).toHaveLength(1);
    expect(plan.summary).toContain('1 write');
    expect(ui!.container.querySelectorAll('.plan-approval .plan-effect').length).toBe(1);

    // The dry-run resolved the pipeline over seeded data: the previewed write carries the COMPUTED
    // literal (550), not the raw `($east | sum revenue)` formula — pure compute fed the effect.
    const planCmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(planCmd).toContain('set Summary!B2');
    expect(planCmd).toContain('550');

    // The host is UNMUTATED before approval: pure ran, but the effect is held at the gate.
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1] ?? '').toBe('');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Only AFTER the one approval does the computed value land in the host (the bridge actuated).
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1]).toBe('550');
  });

  it('a pure-only turn never stages a plan and never touches the bridge', async () => {
    sim = installFakeExcel(paritySeed());
    // A turn of PURE pipelines only (read + filter + sum, a `let`, another scalar). No effect verb,
    // so the orchestrator builds no plan: no `pendingPlan`, no `plan-preview` step, no host mutation.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n' +
          'read Sales!A1:D4 | filter region=East | sum revenue\n' +
          'let $w = read Sales!A1:D4 | filter region=West\n' +
          '$w | count\n' +
          '```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('explore the data, write nothing');
    });
    // It runs to completion WITHOUT ever staging a plan (pure transforms gate on nothing).
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // No plan was ever staged.
    expect(ui!.controller.getState().pendingPlan).toBeUndefined();
    // No plan-preview narration appeared (pure pipelines produce no plan slot).
    const stepKinds = ui!.controller.getState().steps.map((s) => s.kind);
    expect(stepKinds).not.toContain('plan-preview');
    expect(stepKinds).not.toContain('write-result');
    expect(ui!.container.querySelector('.plan-approval')).toBeNull();

    // The bridge was never actuated: every seeded cell is exactly as seeded (no growth, no writes).
    const sales = sim.snapshot().sheets.find((s) => s.name === 'Sales');
    expect(sales?.values).toEqual([
      ['region', 'rep', 'revenue', 'cost'],
      ['East', 'Alice', '300', '120'],
      ['East', 'Bob', '250', '100'],
      ['West', 'Carol', '180', '90'],
    ]);
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1] ?? '').toBe('');
  });

  it('an UNKNOWN verb is rejected by the grammar — never an effect, never reaches the bridge', async () => {
    sim = installFakeExcel(paritySeed());
    // `destroy` is not in the grammar: `parseProgramBlock` yields a `CommandParseError` (did-you-mean)
    // for it, so it can never become a plan slot or reach `bridge.actuate`. The valid `set` in the
    // same block still forms a SINGLE-effect plan.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n' + 'destroy Summary!B2\n' + 'set Summary!B2 42\n' + '```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('clear then set');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The malformed line never became an effect: the plan holds ONLY the valid `set`.
    const plan = ui!.controller.getState().pendingPlan!;
    expect(plan.effects).toHaveLength(1);
    expect(plan.effects[0]!.command).toContain('set Summary!B2');
    expect(ui!.container.querySelectorAll('.plan-approval .plan-effect').length).toBe(1);

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Only the valid effect actuated; the grammar-rejected verb produced no host mutation of its own.
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1]).toBe('42');
    // No comments / formats / extra cells — the unknown verb reached no bridge actuation path.
    expect(sim.snapshot().comments.length).toBe(0);
    expect(sim.snapshot().formats.size).toBe(0);
  });

  it('a misspelled effect verb degrades to a did-you-mean corrective, not a write', async () => {
    sim = installFakeExcel(paritySeed());
    // `st` is a typo of `set`. The grammar rejects it (`unknown verb "st" — did you mean "set"?`)
    // rather than treating it as a write — so no plan slot, no actuation. With NO other effect in the
    // block, the turn stages no plan at all and the host is untouched.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nst Summary!B2 99\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('try to set with a typo');
    });
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The grammar rejected the verb: nothing was ever staged for approval and nothing actuated.
    expect(ui!.controller.getState().pendingPlan).toBeUndefined();
    expect(ui!.container.querySelector('.plan-approval')).toBeNull();
    expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1] ?? '').toBe('');
  });

  it('a non-scalar (table) effect value is rejected by the orchestrator dry-run, never actuated', async () => {
    sim = installFakeExcel(paritySeed());
    // The first effect resolves to a TABLE (`$east` is a filtered table — no scalar terminal), which
    // is not a valid single-cell write. The orchestrator's dry-run rejects it (a corrective `{error}`),
    // so it forms NO plan slot and never reaches the bridge. The SECOND effect (a proper scalar sum)
    // resolves fine — proving the pure pipeline that fed it still composed free.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n' +
          'let $east = read Sales!A1:D4 | filter region=East\n' +
          'set Summary!B2 = ($east)\n' +
          'set Summary!B2 = ($east | sum revenue)\n' +
          '```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('summarize east revenue');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // Only the well-formed scalar effect entered the plan; the table-valued one was rejected at
    // dry-run (never a plan slot, never a bridge call).
    const plan = ui!.controller.getState().pendingPlan!;
    expect(plan.effects).toHaveLength(1);
    const planCmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(planCmd).toContain('550');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The host holds the single computed scalar — never a degenerate multi-row table dump.
    const b2 = sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1] ?? '';
    expect(b2).toBe('550');
    expect(b2).not.toContain('|'); // not a GFM table smuggled into one cell
  });
});
