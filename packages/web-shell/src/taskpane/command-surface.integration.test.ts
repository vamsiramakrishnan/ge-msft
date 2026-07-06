// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { quickActionsForSurface } from '@ge/contracts';
import { quickActionSeed } from './components/quick-action-seed.js';
import {
  installFakeExcel,
  installFakeWord,
  excelSeed,
  scriptedClient,
  mountStack,
  type ExcelSimulator,
  type WordSimulator,
  type MountedStack,
} from '../test-harness/index.js';
import { inferImplicitIntent } from './components/App.js';

/**
 * FULL-STACK interplay for the command surface. Each test installs an in-memory Word host, then
 * drives the REAL stack — `selectBridge('word')` → real `AssistSession` (scripted client) → real
 * `PanelController` → rendered `<App/>` (so `QuickActionBar` + `Composer` are the real components) —
 * and asserts that a quick-action chip / a `/verb` submit routes to the correct controller path:
 *   - read-only `chat` actions and plain questions → `send` (a grounded turn, no gate);
 *   - `write`/`annotation` actions and actuating `/verbs` → `runCommands` → the fail-closed plan gate.
 * No new actuation path is introduced by the chips/palette; they only seed the existing routes.
 */

let sim: WordSimulator | ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

function actionButton(actionId: string): HTMLButtonElement {
  const el = ui!.container.querySelector<HTMLButtonElement>(`button[data-action-id="${actionId}"]`);
  if (!el) throw new Error(`no action button with data-action-id "${actionId}"`);
  return el;
}

function actionTab(label: string): HTMLButtonElement {
  const el = [...ui!.container.querySelectorAll<HTMLButtonElement>('button.action-tab')].find(
    (button) => button.textContent?.includes(label),
  );
  if (!el) throw new Error(`no action tab "${label}"`);
  return el;
}

async function typeAndSubmit(text: string): Promise<void> {
  const input = ui!.container.querySelector<HTMLTextAreaElement>('textarea#ask')!;
  await ui!.act(() => {
    const setter = Object.getOwnPropertyDescriptor(
      window.HTMLTextAreaElement.prototype,
      'value',
    )!.set!;
    setter.call(input, text);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
  await ui!.act(() => {
    ui!.container.querySelector<HTMLButtonElement>('button.snd[type="submit"]')!.click();
  });
}

describe('command surface — quick actions (full-stack)', () => {
  it('a read-only chat chip seeds the exact @-grounded query and renders a grounded answer', async () => {
    sim = installFakeWord();
    const sc = scriptedClient([{ text: 'The contracted SLA is 99.5% monthly.' }]);
    ui = mountStack({ surface: 'word', client: sc });
    await ui.flush();

    const summarize = quickActionsForSurface('word').find((a) => a.id === 'summarize-this')!;
    expect(summarize.output).toBe('chat');
    await ui.act(() => actionTab('Ask').click());
    await ui.act(() => actionButton('summarize-this').click());
    await ui.waitFor(() => (ui!.container.textContent ?? '').includes('99.5% monthly'));

    // It went through `send` with the exact deterministic seed the typed action compiles to
    // (the chip is the same typed Invocation a composer line is — one unified path, no gate staged).
    expect(sc.queries[0]).toBe(quickActionSeed(summarize));
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

  it('a write chip routes through runCommands and stages the fail-closed plan gate', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is 99.9%."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    const tighten = quickActionsForSurface('word').find((a) => a.id === 'tighten')!;
    expect(tighten.output).toBe('write');
    expect(tighten.intent).toBe('rewrite');
    await ui.act(() => actionButton('tighten').click());
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    // The write surfaced as a plan-approval card (the gate) rather than auto-applying.
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
    expect((sim as WordSimulator).snapshot().inserts.length).toBe(0); // nothing applied before approval
  });

  it('an Excel chart chip routes through runCommands and stages an insert-chart gate', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nchart pie Sales!A1:C7 title="Sales mix"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    const createChart = quickActionsForSurface('excel').find((a) => a.id === 'create-chart')!;
    expect(createChart.output).toBe('write');
    expect(createChart.intent).toBe('visualize');
    await ui.act(() => actionButton('create-chart').click());
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    expect(ui.controller.getState().pendingPlan?.effects[0]?.request.kind).toBe('insert-chart');
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
  });
});

describe('command surface — composer / and @ (full-stack)', () => {
  it('a plain @-mention question routes to a grounded send, verbatim', async () => {
    sim = installFakeWord();
    const sc = scriptedClient([{ text: 'It is 99.5% per section 4.' }]);
    ui = mountStack({ surface: 'word', client: sc });
    await ui.flush();

    await typeAndSubmit('@this what is the SLA?');
    await ui.waitFor(() => (ui!.container.textContent ?? '').includes('99.5% per section 4'));

    expect(sc.queries[0]).toBe('@this what is the SLA?');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

  it('a /review verb routes through runCommands and stages the plan gate', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is 99.5% (source needed)."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('/review check the SLA claim');
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
  });

  it('an imperative Excel chart request is promoted to the gated visualize command path', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```cmd\nchart pie Sales!A1:C7 title="Sales mix"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('create a pie chart from Sales!A1:C7');
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    expect(ui.controller.getState().pendingPlan?.effects[0]?.request.kind).toBe('insert-chart');
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
  });

  it('a pasted Excel CLI program routes directly to the gate and writes after approval without a model echo', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Daily schedule',
            origin: 'A1',
            values: Array.from({ length: 12 }, () => Array.from({ length: 9 }, () => '')),
          },
        ],
        activeSheet: 'Daily schedule',
        selection: `'Daily schedule'!B2:I12`,
      }),
    );
    const sc = scriptedClient([]);
    ui = mountStack({ surface: 'excel', client: sc });
    await ui.flush();

    await typeAndSubmit(`
Populate a mock schedule for this please
set 'Daily schedule'!B2 "Time"
set 'Daily schedule'!C2 "Monday"
set 'Daily schedule'!D2 "Tuesday"
set 'Daily schedule'!E2 "Wednesday"
set 'Daily schedule'!F2 "Thursday"
set 'Daily schedule'!G2 "Friday"
set 'Daily schedule'!H2 "Saturday"
set 'Daily schedule'!I2 "Sunday"
set 'Daily schedule'!B3 "08:00 AM"
set 'Daily schedule'!G12 "Wrap Up & Planning"
/summarize @this Summarize the selected range.
`);
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    expect(sc.queries).toEqual([]);
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
    await ui.act(() => ui!.controller.approvePlan());
    await ui.waitFor((s) => s.pendingPlan === undefined && !s.busy);

    const sheet = (sim as ExcelSimulator)
      .snapshot()
      .sheets.find((s) => s.name === 'Daily schedule')!;
    expect(sheet.values[1]?.[1]).toBe('Time');
    expect(sheet.values[1]?.[2]).toBe('Monday');
    expect(sheet.values[1]?.[8]).toBe('Sunday');
    expect(sheet.values[2]?.[1]).toBe('08:00 AM');
    expect(sheet.values[11]?.[6]).toBe('Wrap Up & Planning');
  });
});

describe('command surface — implicit intent inference', () => {
  const base = {
    scope: { kind: 'selection' as const },
    mentions: [],
    instruction: '',
  };

  it('promotes only imperative Excel chart creation when visualize is allowed', () => {
    expect(
      inferImplicitIntent('excel', ['visualize'], {
        ...base,
        raw: 'create a chart from A1:B8',
        instruction: 'create a chart from A1:B8',
      }),
    ).toBe('visualize');
    expect(
      inferImplicitIntent('excel', ['visualize'], {
        ...base,
        raw: 'why are charts not being created?',
        instruction: 'why are charts not being created?',
      }),
    ).toBeUndefined();
    expect(
      inferImplicitIntent('excel', ['ask'], {
        ...base,
        raw: 'create a chart from A1:B8',
        instruction: 'create a chart from A1:B8',
      }),
    ).toBeUndefined();
    expect(
      inferImplicitIntent('word', undefined, {
        ...base,
        raw: 'create a chart from A1:B8',
        instruction: 'create a chart from A1:B8',
      }),
    ).toBeUndefined();
  });

  it('promotes only clear imperative write requests on Word, PowerPoint, and Outlook', () => {
    expect(
      inferImplicitIntent('word', ['rewrite', 'review'], {
        ...base,
        raw: 'rewrite the selected text to be more direct',
        instruction: 'rewrite the selected text to be more direct',
      }),
    ).toBe('rewrite');
    expect(
      inferImplicitIntent('word', ['rewrite', 'review'], {
        ...base,
        raw: 'flag claims that need comments',
        instruction: 'flag claims that need comments',
      }),
    ).toBe('review');
    expect(
      inferImplicitIntent('word', ['ask'], {
        ...base,
        raw: 'rewrite the selected text',
        instruction: 'rewrite the selected text',
      }),
    ).toBeUndefined();
    expect(
      inferImplicitIntent('powerpoint', ['draft'], {
        ...base,
        raw: 'add a slide about Q4 outlook',
        instruction: 'add a slide about Q4 outlook',
      }),
    ).toBe('draft');
    expect(
      inferImplicitIntent('outlook', ['draft'], {
        ...base,
        raw: 'draft a reply to this customer',
        instruction: 'draft a reply to this customer',
      }),
    ).toBe('draft');
    expect(
      inferImplicitIntent('outlook', ['draft'], {
        ...base,
        raw: 'why is draft reply not working?',
        instruction: 'why is draft reply not working?',
      }),
    ).toBeUndefined();
  });
});

describe('command surface — planner-confirm for complex free-text (full-stack, §F)', () => {
  it('a complex /rewrite proposes a plan, then runs the executor on confirm', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        // 1) the PLANNER turn → a ```plan block (the confirmable intention)
        '```plan\nintent rewrite\nsurface word\ncontext inline-preferred\nstep rewrite the SLA to 99.9% as a tracked change\nexclude the indemnity clause\n```',
        // 2) the EXECUTOR turn (after confirm) → a ```cmd suggest → stages the effect gate
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is 99.9%."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    // A composer-typed actuating instruction WITH a constraint → planner-confirm front door.
    await typeAndSubmit('/rewrite the SLA to 99.9% but leave the indemnity clause');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    // The planner card shows the steps; nothing has actuated and no effect gate yet.
    const card = ui.container.querySelector('.command-plan');
    expect(card).not.toBeNull();
    expect(card!.textContent).toContain('99.9%');
    expect(card!.textContent).toContain('Context strategy');
    expect(card!.textContent).toContain('Inline context');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
    expect((sim as WordSimulator).snapshot().inserts.length).toBe(0);

    // Confirm → the executor runs and stages ITS OWN effect-level gate.
    await ui.act(() =>
      ui!.container
        .querySelector<HTMLButtonElement>('[data-testid="command-plan-confirm"]')!
        .click(),
    );
    await ui.waitFor((s) => s.pendingPlan !== undefined);
    expect(ui.controller.getState().pendingCommandPlan).toBeUndefined();
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
    expect((sim as WordSimulator).snapshot().inserts.length).toBe(0); // still nothing applied — gated
  });

  it('a SIMPLE /rewrite skips the planner and goes straight to the executor gate', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is 99.9%."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('/rewrite make it formal');
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    // No planner card — the short instruction went straight to the executor's effect gate.
    expect(ui.container.querySelector('.command-plan')).toBeNull();
    expect(ui.controller.getState().pendingCommandPlan).toBeUndefined();
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
  });
});
