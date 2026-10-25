// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { TriggerRegistry } from '@ge/triggers';
import {
  installFakeExcel,
  excelSeed,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type MountedStack,
  openPlanEffects,
} from '../test-harness/index.js';

/**
 * FULL-STACK Excel integration tests. Each installs an in-memory Excel host (seeded workbook), then
 * drives the REAL stack — `selectBridge('excel')` → real `AssistSession` (scripted fake client) →
 * real `PanelController` → rendered `<App/>` — and asserts both the rendered DOM (UI shows the
 * SIMULATED data) and the mutated fake host (an approved write lands in the workbook). No mocks
 * except the host object model + the model stream.
 *
 * The command loop suspends on async stream consumption between gates, so we `waitFor` a staged
 * gate (`pendingPlan`) before approving, and `waitFor(!busy)` for the run to settle — never a fixed
 * number of microtasks.
 */

let sim: ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

describe('Excel full-stack integration', () => {
  it('refreshContext surfaces the seeded selection + used range as context chips', async () => {
    sim = installFakeExcel();
    ui = mountStack({ surface: 'excel', client: scriptedClient([]) });
    await ui!.flush();
    await ui!.act(() => void ui!.controller.refreshContext());
    await ui!.flush();

    await ui!.showSources();
    const chips = ui!.container.querySelectorAll('.smart-chip');
    expect(chips.length).toBeGreaterThanOrEqual(2);
    const text = ui!.container.textContent ?? '';
    // The seeded selection (Sales!A2:D2) and the used range (Sales!A1:D7) both appear as chips.
    expect(text).toContain('Sales!A2:D2');
    expect(text).toContain('Sales!A1:D7');
  });

  it('a composed read (read | filter | sum) computes the right Value over the seeded workbook', async () => {
    sim = installFakeExcel();
    // Bind East revenue, then write its sum into Summary!B2. The runtime DRY-RUNS the pipeline over
    // the seeded values, so the previewed write carries the computed total (300 + 250 = 550).
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nlet $east = read Sales!A1:D7 | filter region=East\nset Summary!B2 = ($east | sum revenue)\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total East revenue into the summary');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The plan-approval card shows the verbatim effect with the COMPUTED literal, not the pipeline.
    openPlanEffects(ui!.container);
    const planCmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(planCmd).toContain('set Summary!B2');
    expect(planCmd).toContain('550');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Approving mutates the fake host: Summary!B2 holds the computed total.
    const summary = sim.snapshot().sheets.find((s) => s.name === 'Summary');
    expect(summary?.values[1]?.[1]).toBe('550');
  });

  it('a composed read over a NAMED range resolves through getItemOrNullObject + load/sync', async () => {
    sim = installFakeExcel();
    // `SalesTable` (a seeded named range = Sales!A1:D7) is read via the named-range path
    // (names.getItemOrNullObject(name).getRange()), which the fake must register so its load/sync
    // resolves — otherwise reading the range metadata would throw. Same East sum (300 + 250 = 550).
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nlet $east = read SalesTable | filter region=East\nset Summary!B2 = ($east | sum revenue)\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total East revenue from the named table');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    openPlanEffects(ui!.container);
    expect(ui!.container.querySelector('.plan-approval .cmd')?.textContent).toContain('550');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const summary = sim.snapshot().sheets.find((s) => s.name === 'Summary');
    expect(summary?.values[1]?.[1]).toBe('550');
  });

  it('a plan with set + comment renders the real effect-set; approving mutates the host', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Sales!C8 =SUM(C2:C7)\ncomment Sales!C8 "auto-totalled revenue"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total revenue and annotate it');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // Plan card lists BOTH effects with their verbatim command lines.
    const plan = ui!.container.querySelector('.plan-approval');
    expect(plan).not.toBeNull();
    expect(plan?.querySelectorAll('.plan-effect').length).toBe(2);
    openPlanEffects(ui!.container);
    const planText = plan?.textContent ?? '';
    expect(planText).toContain('set Sales!C8 =SUM(C2:C7)');
    expect(planText).toContain('comment Sales!C8');
    // The title counts the effect-set (1 write + 1 comment).
    expect(plan?.querySelector('.plan-approval-title')?.textContent).toBe('Review 2 changes');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const snap = sim.snapshot();
    const sales = snap.sheets.find((s) => s.name === 'Sales');
    // The formula write landed verbatim into the cell (we record `.formulas`, not evaluate it).
    expect(sales?.values[7]?.[2]).toBe('=SUM(C2:C7)');
    // The comment landed on the cell.
    expect(
      snap.comments.some((c) => c.cell.includes('C8') && c.content.includes('auto-totalled')),
    ).toBe(true);
  });

  it('rejecting the plan writes nothing (per-plan gating holds)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nset Sales!C8 =SUM(C2:C7)\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total revenue');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    expect(ui!.container.querySelector('.plan-approval')).not.toBeNull();

    await ui!.act(() => ui!.controller.rejectPlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Nothing actuated: the cell is still empty (outside the originally-seeded grid).
    const sales = sim.snapshot().sheets.find((s) => s.name === 'Sales');
    expect(sales?.values[7]?.[2] ?? '').toBe('');
    // The plan card is gone after the decision.
    expect(ui!.container.querySelector('.plan-approval')).toBeNull();
  });

  it('fail-closed: a write blocked by the actuation gate never mutates the host', async () => {
    sim = installFakeExcel();
    const gate = new TriggerRegistry();
    gate.register({
      id: 'veto',
      on: 'pre-actuation',
      handle: () => ({ kind: 'block', reason: 'policy: no edits to Sales' }),
    });
    ui = mountStack({
      surface: 'excel',
      triggers: gate,
      client: scriptedClient(['```cmd\nset Sales!C8 =SUM(C2:C7)\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total revenue');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    // Even with plan approval, the gate blocks the actual actuation.
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const sales = sim.snapshot().sheets.find((s) => s.name === 'Sales');
    expect(sales?.values[7]?.[2] ?? '').toBe('');
    // The blocked write narrated as a run-step.
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps.toLowerCase()).toContain('blocked');
  });

  it('an approved write persists provenance into the fake settings bag', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nset Summary!B2 42\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write a value');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The reversible write recorded a durable provenance record (keyed by changeId) and saved it.
    expect(sim.office.settings.size).toBeGreaterThan(0);
    expect(sim.office.settingsSaved).toBe(true);
    const record = [...sim.office.settings.values()][0];
    expect(String(record)).toContain('sim.user@acme');
  });

  it('an unsupported actuation on a stale host degrades rather than crashing (comments < 1.10)', async () => {
    // Seed a host BELOW ExcelApi 1.10 so add-comment is unsupported; the loop degrades.
    sim = installFakeExcel(undefined, { ExcelApi: 9 });
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\ncomment Sales!A2 "needs review"\n```', '```cmd\ndone\n```'])
        .client,
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag a cell');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // No comment was added (host too old); the loop degraded without throwing.
    expect(sim.snapshot().comments.length).toBe(0);
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps.length).toBeGreaterThan(0);
  });

  it('sorts a range natively, verifies it, and Undo restores the earlier order (WS4b)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n/sort-range range=Sales!A1:D7 key=revenue order=descending\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('sort by revenue, highest first');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const revenue = () =>
      sim!
        .snapshot()
        .sheets[0]!.values.slice(1)
        .map((row) => row[2]);
    expect(revenue()).toEqual(['300', '250', '220', '180', '140', '90']);
    expect(sim.snapshot().sheets[0]!.values[0]).toEqual(['region', 'rep', 'revenue', 'cost']);
    const reply = ui!.controller.getState().messages.at(-1)!;
    expect(reply.changes?.[0]?.undoable).toBe(true);

    let undo!: Promise<void>;
    await ui!.act(() => {
      undo = ui!.controller.undoChanges(reply.id);
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await undo;
    expect(sim.snapshot().sheets[0]!.values).toEqual(before);
  });

  it.each([
    ['a range narrower than its table', '/sort-range range=Sales!A1:C7 key=revenue'],
    ['a key name when headers=false', '/sort-range range=Sales!A1:D7 key=revenue headers=false'],
  ])('refuses %s, changing nothing (security review M3, L4)', async (_label, cmd) => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([`\`\`\`cmd\n${cmd}\n\`\`\``, '```cmd\ndone "No."\n```']),
    });
    await ui!.flush();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('sort it');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined || !s.busy);
    if (ui!.controller.getState().pendingPlan) await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    expect(sim.snapshot().sheets[0]!.values).toEqual(before);
  });

  it('a sort of rows already in order succeeds without moving them, even over formulas', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Summary',
            origin: 'A1',
            values: [
              ['branch', 'growth'],
              ['Medan', '=18.25'],
              ['Surabaya', '=16.22'],
              ['Jakarta', '=12.17'],
            ],
          },
        ],
      }),
    );
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n/sort-range range=Summary!A1:B4 key=growth order=descending\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('sort by growth, highest first');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    expect(sim.snapshot().sheets[0]!.values).toEqual(before);
    const steps = ui!.controller
      .getState()
      .steps.map((s) => s.text ?? '')
      .join('\n');
    expect(steps).not.toContain('has_formulas');
    expect(ui!.controller.getState().messages.at(-1)!.changes?.[0]?.kind).toBe('sort-range');
  });

  // One batch, run in order: a later effect sees the cells an earlier effect of the SAME batch
  // wrote. That is not a stale target. A change from outside the batch still is.
  async function runBatch(cmd: string, beforeApprove?: () => void): Promise<string> {
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        `\`\`\`cmd\n${cmd}\n\`\`\``,
        '```cmd\ndone\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('do it');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    beforeApprove?.();
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();
    return ui!.controller
      .getState()
      .steps.map((s) => s.text ?? '')
      .join('\n');
  }

  it('a write over cells an earlier write of the same batch wrote lands', async () => {
    sim = installFakeExcel();
    const steps = await runBatch('set Sales!C2 999\nset Sales!C2 1000');
    expect(steps).not.toContain('stale_target');
    expect(sim.snapshot().sheets[0]!.values[1]![2]).toBe('1000');
  });

  it('a sort after a write of the same batch sorts the written rows', async () => {
    sim = installFakeExcel();
    const steps = await runBatch(
      'set Sales!C2 999\n/sort-range range=Sales!A1:D7 key=revenue order=descending',
    );
    expect(steps).not.toContain('stale_target');
    const revenue = sim
      .snapshot()
      .sheets[0]!.values.slice(1)
      .map((row) => row[2]);
    expect(revenue).toEqual(['999', '250', '220', '180', '140', '90']);
  });

  it('a sort after a same-row formula column of the same batch sorts the rows', async () => {
    sim = installFakeExcel();
    const steps = await runBatch(
      'set Sales!E1 "ratio"\nset Sales!E2:E7 =C2/D2\n/sort-range range=Sales!A1:E7 key=revenue order=descending',
    );
    expect(steps).not.toMatch(/stale_target|has_formulas/);
    const revenue = sim
      .snapshot()
      .sheets[0]!.values.slice(1)
      .map((row) => row[2]);
    expect(revenue).toEqual(['300', '250', '220', '180', '140', '90']);
  });

  it('a sort over a formula that uses another row is still refused', async () => {
    sim = installFakeExcel();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);
    const steps = await runBatch(
      'set Sales!E1 "running"\nset Sales!E2:E7 =C2+E1\n/sort-range range=Sales!A1:E7 key=revenue order=descending',
    );
    expect(steps).toContain('has_formulas');
    expect(sim.snapshot().sheets[0]!.values.map((row) => row.slice(0, 4))).toEqual(before);
  });

  it('a sheet named with a stray leading space is found when the command drops the space', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          { name: 'Branch Performance', origin: 'A1', values: [['branch'], ['Medan']] },
          { name: ' Ringkasan Cabang', origin: 'A1', values: [['']] },
        ],
      }),
    );
    const steps = await runBatch(`grid 'Ringkasan Cabang'!A1:B2 = "Cabang\\tTotal\\nMedan\\t10"`);
    expect(steps).not.toMatch(/prepare_failed|prerequisite_failed/);
    const summary = sim.snapshot().sheets.find((s) => s.name === ' Ringkasan Cabang')!;
    expect(summary.values[1]).toEqual(['Medan', '10']);
  });

  it('a change outside the batch after preview still stops the sort', async () => {
    sim = installFakeExcel();
    const steps = await runBatch(
      'set Sales!C2 999\n/sort-range range=Sales!A1:D7 key=revenue order=descending',
      () => {
        // The user edits a cell the batch does not write, between preview and Apply.
        sim!.seed.sheets[0]!.values[3]![3] = '5';
      },
    );
    expect(steps).toContain('stale_target');
    expect(sim.snapshot().sheets[0]!.values[1]![0]).toBe('East');
  });

  it('a format of the range does not hide an outside edit from the sort (security review H1)', async () => {
    sim = installFakeExcel();
    const steps = await runBatch(
      'format Sales!A1:D7 bold=true\n/sort-range range=Sales!A1:D7 key=revenue order=descending',
      () => {
        sim!.seed.sheets[0]!.values[3]![3] = '5';
      },
    );
    expect(steps).toContain('stale_target');
    expect(sim.snapshot().sheets[0]!.values[1]![0]).toBe('East');
  });

  it('one Undo restores two overlapping changes of the same batch (security review M3)', async () => {
    sim = installFakeExcel();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);
    await runBatch('set Sales!C2 999\nset Sales!C2 1000');
    expect(sim.snapshot().sheets[0]!.values[1]![2]).toBe('1000');
    const reply = ui!.controller.getState().messages.find((m) => (m.changes?.length ?? 0) === 2)!;
    let undo!: Promise<void>;
    await ui!.act(() => {
      undo = ui!.controller.undoChanges(reply.id);
    });
    for (let i = 0; i < 2; i++) {
      await ui!.waitFor((s) => s.pendingPlan !== undefined);
      await ui!.act(() => ui!.controller.approvePlan());
      await ui!.waitFor((s) => s.pendingPlan === undefined);
    }
    await ui!.waitFor((s) => !s.busy);
    await undo;
    expect(sim.snapshot().sheets[0]!.values).toEqual(before);
  });

  it('a sort key that is not a column is refused before anything changes (WS4b)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\n/sort-range range=Sales!A1:D7 key=Deposits\n```',
        '```cmd\ndone "There is no Deposits column."\n```',
      ]),
    });
    await ui!.flush();
    const before = sim.snapshot().sheets[0]!.values.map((row) => [...row]);
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('sort by deposits');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined || !s.busy);
    if (ui!.controller.getState().pendingPlan) await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    expect(sim.snapshot().sheets[0]!.values).toEqual(before);
  });

  it('a thick outside and thin inside border land on the right edges (WS4b)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nformat Sales!A1:D7 borderOutside=thick borderInside=thin\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('thick outside border, thin inside');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    const borders = sim.snapshot().formats.get('Sales!A1:D7')?.borders ?? {};
    for (const edge of ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'])
      expect(borders[edge]?.weight, edge).toBe('Thick');
    for (const edge of ['InsideHorizontal', 'InsideVertical'])
      expect(borders[edge]?.weight, edge).toBe('Thin');
  });

  it('an approved format effect records the facets against the targeted range', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nformat Sales!A1:D1 bold=true fill=#FFF2CC\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('bold the header row');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The format-cells write landed against the header range (write-only on a Range, recorded apart
    // from the value grid).
    const fmt = sim.snapshot().formats.get('Sales!A1:D1');
    expect(fmt?.bold).toBe(true);
    expect(fmt?.fill).toBe('#FFF2CC');
  });

  it('a range set repeated in a later turn is not applied (or approved) twice (UAT #7c)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Sales!E1 "Double"\nset Sales!E2:E4 =C2*2\nformat Sales!E2:E4 numberFormat="0.0"\n```',
        '```cmd\nfinish when=verified\ndone "Filled."\n```',
        "```cmd\nread 'Sales'!A1:E4\n```",
        // The model quotes the sheet this time; the recorded write has no quotes.
        '```cmd\nset \'Sales\'!E2:E4 =C2*2\nformat \'Sales\'!$E$2:$E$4 numberFormat="0.0"\ndone "Filled E2:E4."\n```',
      ]),
    });
    await ui!.flush();

    let approvals = 0;
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('double revenue into E');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    approvals++;
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy || s.pendingPlan !== undefined);
    if (ui!.controller.getState().pendingPlan) {
      approvals++;
      await ui!.act(() => ui!.controller.approvePlan());
    }
    await ui!.waitFor((s) => !s.busy);
    await run;

    expect(approvals).toBe(1);
    expect(sim.snapshot().sheets.find((s) => s.name === 'Sales')?.values[3]?.[4]).toBe('=C4*2');
  });

  it('a verified finish after a cell write + format ends done, not "did not finish" (UAT #10)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Sales!C8 =SUM(C2:C7)\nformat Sales!C8 numberFormat="0.0%"\nfinish when=verified\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('total revenue as a percentage');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    expect(sim.snapshot().formats.get('Sales!C8')?.numberFormat).toBe('0.0%');
    // The format has no readback; that is not a failure. Only the result card shows.
    expect(ui!.container.querySelector('.error-card')).toBeNull();
    expect(ui!.container.querySelector('.result-card')).not.toBeNull();
  });

  it('chart by= builds the summary table from the cells read, then charts it — one approval (UAT #11)', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nchart column Sales!A1:D7 x=rep by=region value=revenue title="Revenue"\n```',
        '```cmd\ndone "Charted."\n```',
      ]),
    });
    await ui!.flush();
    let approvals = 0;
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('chart revenue per region by rep');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    approvals++;
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;

    expect(approvals).toBe(1);
    const sales = sim.snapshot().sheets.find((s) => s.name === 'Sales')!;
    // Summary at F1: reps down, regions across; every cell a live exact-match sum, never a typed number.
    expect(sales.values[0]?.slice(5, 10)).toEqual(['rep', 'East', 'West', 'North', 'South']);
    expect(sales.values[1]?.[5]).toBe('Alice');
    expect(sales.values[1]?.[6]).toBe('=SUMPRODUCT(--($A$2:$A$7=G$1),--($B$2:$B$7=$F2),$C$2:$C$7)');
    const chart = sim.snapshot().charts.find((c) => c.title === 'Revenue');
    expect(chart?.source).toContain('F1:J7');
  });

  it('chart by= over month text keeps its labels as text, so the totals match (real Excel, 2026-10-17)', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Branch',
            origin: 'A1',
            values: [
              ['Branch', 'Month', 'Complaints'],
              ['Jakarta', 'Jan 2026', '4'],
              ['Jakarta', 'Feb 2026', '5'],
              ['Medan', 'Jan 2026', '6'],
              ['Medan', 'Feb 2026', '5'],
            ],
          },
        ],
      }),
    );
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nchart line Branch!A1:C5 x=Month by=Branch value=Complaints\n```',
        '```cmd\ndone "Charted."\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('trend of complaints per branch');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;

    const sheet = sim.snapshot().sheets.find((s) => s.name === 'Branch')!;
    // Before the fix the simulated Excel stored these labels as date serials (46023), as real Excel did.
    expect(sheet.values[1]?.[4]).toBe('Jan 2026');
    expect(sheet.values[2]?.[4]).toBe('Feb 2026');
    expect(sim.snapshot().charts).toHaveLength(1);
  });

  it('a retried chart by= rewrites its own summary block, never a second copy (live 2026-10-17)', async () => {
    sim = installFakeExcel();
    const chart =
      '```cmd\nchart column Sales!A1:D7 x=rep by=region value=revenue title="Revenue"\n```';
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([chart, chart, '```cmd\ndone "Charted."\n```']),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('chart revenue per region by rep');
    });
    for (let i = 0; i < 2; i++) {
      await ui!.waitFor((s) => s.pendingPlan !== undefined || !s.busy);
      if (!ui!.controller.getState().pendingPlan) break;
      await ui!.act(() => ui!.controller.approvePlan());
    }
    await ui!.waitFor((s) => !s.busy);
    await run;

    const sales = sim.snapshot().sheets.find((s) => s.name === 'Sales')!;
    expect(sales.values[0]?.slice(5, 10)).toEqual(['rep', 'East', 'West', 'North', 'South']);
    // Nothing beside it: the retry did not find "its" block taken and move on.
    expect(sales.values.every((row) => (row[11] ?? '') === '')).toBe(true);
  });
});
