// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { TriggerRegistry } from '@ge/triggers';
import {
  installFakeExcel,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type MountedStack,
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

    const chips = ui!.container.querySelectorAll('.chip');
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
    const planText = plan?.textContent ?? '';
    expect(planText).toContain('set Sales!C8 =SUM(C2:C7)');
    expect(planText).toContain('comment Sales!C8');
    // The summary header counts the effect-set (1 write + 1 comment).
    expect(plan?.querySelector('.pin')?.textContent).toContain('1 write');

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
});
