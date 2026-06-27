// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  installFakeExcel,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * FULL-STACK interplay: a single composed turn that carries MULTIPLE effects → the ADR-0005
 * plan-approval gate. Wires the REAL packages across their seam — `selectBridge('excel')` → real
 * `AssistSession` (the runtime planner/executor) → real `PanelController` → rendered `<App/>` (the
 * web-shell plan-approval card) — over a scripted fake client. Nothing between the host object model
 * and the model stream is mocked, so each assertion is an assertion about the real
 * runtime↔web-shell↔triggers contract:
 *
 *   1. Approving the plan applies EVERY effect (read back from the fake workbook).
 *   2. Rejecting the plan applies NONE of them (fail-closed for the whole set).
 *   3. The gate REPORTS the effect cardinality — the summary header counts the kinds, and the card
 *      renders one row per effect.
 *   4. Pure-composition steps (`let $x = read … | filter …`) in the SAME plan do NOT consume the
 *      per-turn effect budget: a turn whose (pure + effects) entry count EXCEEDS the write cap still
 *      lands every effect, because only effect verbs reserve a plan slot.
 *
 * The command loop suspends on async stream consumption between gates, so we `waitFor` the staged
 * `pendingPlan` before deciding and `waitFor(!busy)` for the run to settle — never a fixed tick count.
 */

let sim: ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

/** Read a single A1 cell's value out of a sheet in the snapshot (0-based row/col within its grid). */
function cell(snap: ReturnType<ExcelSimulator['snapshot']>, sheet: string, r: number, c: number) {
  return snap.sheets.find((s) => s.name === sheet)?.values[r]?.[c] ?? '';
}

describe('multi-effect plan approval (runtime + web-shell + triggers)', () => {
  it('approving a multi-effect plan applies EVERY effect to the host', async () => {
    sim = installFakeExcel();
    // ONE turn, THREE write effects to distinct cells. The runtime composes them into one plan; the
    // gate takes a single approval; the executor then actuates each effect one-by-one.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Summary!B2 100\nset Summary!B3 200\nset Summary!B4 300\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write three summary rows');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The card stages exactly three effects before anything actuates.
    expect(ui!.controller.getState().pendingPlan?.effects.length).toBe(3);
    expect(ui!.container.querySelectorAll('.plan-approval .plan-effect').length).toBe(3);

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // ALL three writes landed in the fake workbook (Summary grid: row 0 header, row 1 = B2, …).
    const snap = sim.snapshot();
    expect(cell(snap, 'Summary', 1, 1)).toBe('100');
    expect(cell(snap, 'Summary', 2, 1)).toBe('200');
    expect(cell(snap, 'Summary', 3, 1)).toBe('300');
  });

  it('rejecting a multi-effect plan applies NONE of the effects', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Summary!B2 100\nset Summary!B3 200\nset Summary!B4 300\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write three summary rows');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    expect(ui!.controller.getState().pendingPlan?.effects.length).toBe(3);

    await ui!.act(() => ui!.controller.rejectPlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The reject blocked the WHOLE plan: not one of the three cells was written.
    const snap = sim.snapshot();
    expect(cell(snap, 'Summary', 1, 1)).toBe(''); // B2 still empty (seeded as '')
    expect(cell(snap, 'Summary', 2, 1)).toBe(''); // B3 outside the seeded grid
    expect(cell(snap, 'Summary', 3, 1)).toBe('');
    // The card is gone after the decision, and no provenance was persisted for a blocked plan.
    expect(ui!.container.querySelector('.plan-approval')).toBeNull();
    expect(sim.office.settings.size).toBe(0);
  });

  it('the gate REPORTS effect cardinality across kinds (writes + comments)', async () => {
    sim = installFakeExcel();
    // Two writes + one comment → the plan summary must count them by kind.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Summary!B2 100\nset Summary!B3 200\ncomment Sales!A2 "flagged for review"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write totals and annotate');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    const plan = ui!.container.querySelector('.plan-approval');
    expect(plan).not.toBeNull();
    // Cardinality, in the DOM and in the controller state.
    expect(plan!.querySelectorAll('.plan-effect').length).toBe(3);
    const pin = plan!.querySelector('.pin')?.textContent ?? '';
    // The summary groups by kind and pluralizes: "2 writes + 1 comment".
    expect(pin).toContain('2 writes');
    expect(pin).toContain('1 comment');
    expect(ui!.controller.getState().pendingPlan?.summary).toBe('2 writes + 1 comment');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Approving the mixed-kind plan actually applied all three effects.
    const snap = sim.snapshot();
    expect(cell(snap, 'Summary', 1, 1)).toBe('100');
    expect(cell(snap, 'Summary', 2, 1)).toBe('200');
    expect(snap.comments.some((c) => c.cell.includes('A2') && c.content.includes('flagged'))).toBe(
      true,
    );
  });

  it('pure-composition steps do NOT consume the per-turn effect budget', async () => {
    sim = installFakeExcel();
    // ONE turn mixing FIVE pure-composition steps (`let $x = read … | filter …`) with SEVEN write
    // effects. The total entry count (12) exceeds the per-turn WRITE cap (default 8), but only the
    // seven effect verbs reserve plan slots — the pure binds are analysis-only and never enter the
    // effect-set. If pure steps consumed the effect budget, some of the seven writes would be capped
    // and dropped; the seam contract is that all seven survive into one plan and all seven actuate.
    const block = [
      '```cmd',
      'let $a = read Sales!A1:D7 | filter region=East',
      'let $b = read Sales!A1:D7 | filter region=West',
      'let $c = read Sales!A1:D7 | filter region=North',
      'let $d = read Sales!A1:D7 | filter region=South',
      'let $e = read Sales!A1:D7 | filter rep=Alice',
      'set Summary!B2 11',
      'set Summary!B3 22',
      'set Summary!B4 33',
      'set Summary!B5 44',
      'set Summary!B6 55',
      'set Summary!B7 66',
      'set Summary!B8 77',
      '```',
    ].join('\n');
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([block, '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('compute regional cuts then write seven rows');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The plan carries exactly the SEVEN effects — the five pure binds did not enter the effect-set,
    // so none of the seven writes was crowded out of the budget.
    const pending = ui!.controller.getState().pendingPlan;
    expect(pending?.effects.length).toBe(7);
    expect(pending?.summary).toBe('7 writes');
    expect(ui!.container.querySelectorAll('.plan-approval .plan-effect').length).toBe(7);

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Every one of the seven writes actuated (none capped) — the budget was spent only by effects.
    const snap = sim.snapshot();
    expect(cell(snap, 'Summary', 1, 1)).toBe('11');
    expect(cell(snap, 'Summary', 2, 1)).toBe('22');
    expect(cell(snap, 'Summary', 3, 1)).toBe('33');
    expect(cell(snap, 'Summary', 4, 1)).toBe('44');
    expect(cell(snap, 'Summary', 5, 1)).toBe('55');
    expect(cell(snap, 'Summary', 6, 1)).toBe('66');
    expect(cell(snap, 'Summary', 7, 1)).toBe('77');

    // None of the write-result steps was capped (the budget guard never tripped for an effect).
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps.toLowerCase()).not.toContain('write cap');
  });
});
