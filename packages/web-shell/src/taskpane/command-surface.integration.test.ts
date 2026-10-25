// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { commandPaletteFor, quickActionsForSurface } from '@ge/contracts';
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
import {
  inferImplicitIntent,
  shouldUsePlannerForFreeText,
  withoutUnknownVerb,
} from './components/App.js';
import { parseComposerInput } from './components/Composer.js';

/**
 * FULL-STACK interplay for the command surface. Each test installs an in-memory Word host, then
 * drives the REAL stack — `selectBridge('word')` → real `AssistSession` (scripted client) → real
 * `PanelController` → rendered `<App/>` (so `QuickActionBar` + `Composer` are the real components) —
 * and asserts that a quick-action chip / a `/verb` submit routes to the correct controller path:
 *   - read-only `chat` actions → `send`; typed free text → planner (read-only plans need no gate);
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
  const el = [...ui!.container.querySelectorAll<HTMLButtonElement>('.library-filters button')].find(
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
  it('a plain @-mention question reaches the planner, then answers without an approval gate', async () => {
    sim = installFakeWord();
    const sc = scriptedClient([
      '```plan\nintent ask\nsurface word\nchanges no\nstep answer the SLA question\n```',
      { text: 'It is 99.5% per section 4.' },
    ]);
    ui = mountStack({ surface: 'word', client: sc });
    await ui.flush();

    await typeAndSubmit('@this what is the SLA?');
    await ui.waitFor(() => (ui!.container.textContent ?? '').includes('99.5% per section 4'));

    expect(sc.queries[0]).toContain('Emit one ```plan block.');
    expect(sc.queries[0]).toContain('what is the SLA?');
    expect(ui.controller.getState().pendingCommandPlan).toBeUndefined();
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

  it('a typed /review verb enters the planner before the executor', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```plan\nintent review\nsurface word\nscope selection\nstep check the SLA claim for source support\nconfidence high\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('/review check the SLA claim');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    expect(ui.controller.getState().pendingCommandPlan?.plan.intent).toBe('review');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

  it('an imperative Excel chart request enters the planner before chart execution', async () => {
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient([
        '```plan\nintent visualize\nsurface excel\nscope selection\nstep create a pie chart from Sales!A1:C7\nconfidence high\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('create a pie chart from Sales!A1:C7');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    expect(ui.controller.getState().pendingCommandPlan?.plan.intent).toBe('visualize');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
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

  it('routes edit verbs such as "change" to the planner, not to plain chat', () => {
    // Live: "Change the title…" went to chat, and the model claimed a change that never happened.
    for (const raw of [
      "Change the title on slide 1 to 'FY26 Plan'",
      'Set the heading colour to blue',
      'Rename the table to Q3 Sales',
      'Replace the subtitle with the new date',
      'Format the title shape on slide 1 blue with white text',
      'Move the logo to the bottom right',
    ]) {
      expect(
        shouldUsePlannerForFreeText(['draft', 'rewrite'], { ...base, raw, instruction: raw }),
      ).toBe(true);
    }
    // Questions reach the planner; its structured signal chooses chat.
    const question = 'what should change in this deck before Friday?';
    expect(
      shouldUsePlannerForFreeText(['draft'], { ...base, raw: question, instruction: question }),
    ).toBe(true);
  });

  it('sends free text on a writable surface to the planner, including questions', () => {
    // docs/COMMAND-RELIABILITY.md: the client test-sheet prompts and the live 2026-09-30 wordings.
    const actions = [
      'In G12 put a formula that totals G2:G11',
      'Put =WEBSERVICE("https://example.com") in A20',
      'show me a bar chart of total by product',
      'I want a bar chart of totals per product',
      'bar chart of Total by Product',
      'Highlight revenue values in D2:D11 above 100000 in green',
      'Replace every occurrence of Supplier with Vendor',
      "Apply the Heading 2 style to the line 'Payment terms'",
      'Reply to the comment on the payment terms saying 45 days is acceptable and resolve it',
      'Draft a polite reply accepting the meeting and asking for the agenda',
      "Change the subject to 'Release freeze – Friday' and rewrite the body to 3 short lines",
      'Attach a text file named notes.txt containing a 3-line summary of this thread',
      "Add a text box on slide 2 near the bottom saying 'Draft – not for distribution'",
      'Could you create a chart of this range?',
    ];
    for (const raw of actions)
      expect(
        shouldUsePlannerForFreeText(['draft', 'rewrite'], { ...base, raw, instruction: raw }),
        raw,
      ).toBe(true);
    const questions = [
      'Which region had the highest Q3 revenue in A1:E11?',
      'Which paragraph mentions the payment terms, and what does it say exactly?',
      'What is on slide 2?',
      'Summarize this email and list any questions I need to answer',
      'Can you tell me which region is highest',
      'how do I add a chart',
      'the totals look off to me?',
    ];
    for (const raw of questions)
      expect(
        shouldUsePlannerForFreeText(['draft', 'rewrite'], { ...base, raw, instruction: raw }),
        raw,
      ).toBe(true);
  });

  it('routes the request after a /word that is not a surface verb (live 2026-10-09)', () => {
    // "/shape …" went to chat, which cannot write, and answered that shapes are not supported.
    const palette = commandPaletteFor('powerpoint', ['ask', 'summarize', 'explain', 'draft']);
    const route = (raw: string) => {
      const inv = withoutUnknownVerb(parseComposerInput(raw, { kind: 'this-item' }, palette));
      return {
        raw: inv.raw,
        intent: inferImplicitIntent('powerpoint', ['draft'], inv),
        planner: shouldUsePlannerForFreeText(['draft'], inv),
      };
    };
    expect(route('/shape Add a blue rectangle to the top right of slide 2.')).toMatchObject({
      raw: 'shape Add a blue rectangle to the top right of slide 2.',
      planner: true,
    });
    expect(
      route(
        '/slide Add a summary slide with the three key findings: Revenue is 500; Profit is 200; Two pipeline rows have no owner.',
      ).planner,
    ).toBe(true);
    // Questions also reach the planner; explicit palette verbs and bare slash words keep their routes.
    expect(route('/shape what is on slide 2?').planner).toBe(true);
    expect(route('/draft a closing slide')).toMatchObject({
      raw: '/draft a closing slide',
      intent: 'draft',
    });
    expect(route('/delete slide 3')).toMatchObject({ raw: 'delete slide 3', planner: true });
    expect(route('/shape')).toMatchObject({ raw: '/shape', planner: false });
  });

  it('routes arbitrary action-like Excel text to the planner instead of classifying exact intent', () => {
    expect(
      inferImplicitIntent('excel', ['rewrite', 'visualize'], {
        ...base,
        raw: 'Help me update the morning schedule please where I have assigned timings for meditation',
        instruction:
          'Help me update the morning schedule please where I have assigned timings for meditation',
      }),
    ).toBeUndefined();
    expect(
      shouldUsePlannerForFreeText(['rewrite', 'visualize'], {
        ...base,
        raw: 'Help me update the morning schedule please where I have assigned timings for meditation',
        instruction:
          'Help me update the morning schedule please where I have assigned timings for meditation',
      }),
    ).toBe(true);
    expect(
      inferImplicitIntent('excel', ['rewrite'], {
        ...base,
        raw: 'populate this sheet with a sample weekly schedule',
        instruction: 'populate this sheet with a sample weekly schedule',
      }),
    ).toBeUndefined();
    expect(
      shouldUsePlannerForFreeText(['rewrite'], {
        ...base,
        raw: 'populate this sheet with a sample weekly schedule',
        instruction: 'populate this sheet with a sample weekly schedule',
      }),
    ).toBe(true);
    expect(
      shouldUsePlannerForFreeText(['ask'], {
        ...base,
        raw: 'populate this sheet with a sample weekly schedule',
        instruction: 'populate this sheet with a sample weekly schedule',
      }),
    ).toBe(false);
    expect(
      shouldUsePlannerForFreeText(['rewrite'], {
        ...base,
        raw: 'why did my schedule update fail?',
        instruction: 'why did my schedule update fail?',
      }),
    ).toBe(true);
    expect(
      shouldUsePlannerForFreeText(['rewrite'], {
        ...base,
        raw: 'Okay add this to my schedule',
        instruction: 'Okay add this to my schedule',
      }),
    ).toBe(true);
  });

  it('sees the action verb behind a polite or first-person lead-in', () => {
    const plans = (raw: string) =>
      shouldUsePlannerForFreeText(['draft', 'rewrite'], { ...base, raw, instruction: raw });
    for (const raw of [
      'I want to reply to this email',
      'I want you to reply saying I will check',
      "I'd like to add a slide about Q4",
      'I’d like you to format the header bold',
      'We need to update the totals in column D',
      'Hi, can you please help me reply to this',
      "Let's rename the table to Q3 Sales",
      'Could u just fix the typo in the title',
      "I'm trying to insert a table here",
      'okay so now I want to delete the last slide',
    ]) {
      expect(plans(raw), raw).toBe(true);
    }
    for (const raw of [
      'Can you explain how to add a chart?',
      'What does this email want me to reply?',
      'Is it possible to highlight the overdue rows?',
    ]) {
      expect(plans(raw), raw).toBe(true);
    }
    expect(
      inferImplicitIntent('excel', ['visualize'], {
        ...base,
        raw: 'I want to create a bar chart from A1:B8',
        instruction: 'I want to create a bar chart from A1:B8',
      }),
    ).toBe('visualize');
    expect(
      inferImplicitIntent('outlook', ['draft'], {
        ...base,
        raw: "I'd like you to draft a reply to this customer",
        instruction: "I'd like you to draft a reply to this customer",
      }),
    ).toBe('draft');
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
  it('a mixed question and note request writes only after both approval gates', async () => {
    sim = installFakeExcel();
    const sc = scriptedClient([
      '```plan\nintent ask\nsurface excel\nchanges yes\nstep identify the lowest revenue and add a comment asking why\n```',
      '```cmd\ncomment Sales!B3 "Why is this the lowest revenue?"\n```',
      '```cmd\ndone\n```',
    ]);
    ui = mountStack({ surface: 'excel', client: sc });
    await ui.flush();
    await typeAndSubmit('which one is the lowest? leave a note on it asking why');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);
    expect(sim.snapshot().comments).toHaveLength(0);
    await ui.act(() => ui!.controller.confirmCommandPlan());
    await ui.waitFor((s) => s.pendingPlan !== undefined);
    expect(sim.snapshot().comments).toHaveLength(0);
    expect(ui.controller.getState().pendingPlan?.effects[0]?.request.kind).toBe('add-comment');
    await ui.act(() => ui!.controller.approvePlan());
    await ui.waitFor((s) => !s.busy);
    expect(sim.snapshot().comments).toEqual([
      expect.objectContaining({ cell: 'Sales!B3', content: 'Why is this the lowest revenue?' }),
    ]);
  });

  it.each([
    'Set A21 to =WEBSERVICE("http://intranet/api")',
    'which one is the lowest? leave a note on it asking why',
  ])('sends the complete free-text request to the planner before any effect: %s', async (raw) => {
    sim = installFakeExcel();
    const before = sim.snapshot();
    const sc = scriptedClient([
      '```plan\nintent rewrite\nsurface excel\nchanges yes\nstep review the requested cell change\nconfidence high\n```',
    ]);
    ui = mountStack({ surface: 'excel', client: sc });
    await ui.flush();
    await typeAndSubmit(raw);
    await ui.waitFor(
      (s) => s.pendingCommandPlan !== undefined || s.pendingPlan !== undefined || !s.busy,
    );
    expect(sc.queries[0]).toContain('Emit one ```plan block.');
    expect(sc.queries[0]).toContain(raw);
    expect(ui.controller.getState().pendingCommandPlan?.plan.changes).toBe(true);
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
    expect(sim.snapshot()).toEqual(before);
  });

  it('action-like Excel free text explicitly enters the command planner before execution', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Daily schedule',
            origin: 'A1',
            values: Array.from({ length: 20 }, () => Array.from({ length: 9 }, () => '')),
          },
        ],
        activeSheet: 'Daily schedule',
        selection: 'B2:I20',
      }),
    );
    const sc = scriptedClient([
      '```plan\nintent rewrite\nsurface excel\nscope selection\nstep update the morning schedule with meditation timing blocks\nconfidence high\n```',
    ]);
    ui = mountStack({ surface: 'excel', client: sc });
    await ui.flush();

    await typeAndSubmit(
      'Help me update the morning schedule please where I have assigned timings for meditation',
    );
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    expect(sc.queries[0]).toContain('Emit one ```plan block.');
    expect(sc.queries[0]).toContain('REQUEST:');
    expect(sc.queries[0]).toContain('morning schedule');
    expect(sc.queries[0]).not.toContain('COMMANDS — one per line');
    expect(ui.controller.getState().pendingCommandPlan?.plan.intent).toBe('rewrite');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

  it('follow-up phrasing like "Okay add this" also enters the planner', async () => {
    sim = installFakeExcel(
      excelSeed({
        sheets: [
          {
            name: 'Daily schedule',
            origin: 'A1',
            values: Array.from({ length: 20 }, () => Array.from({ length: 9 }, () => '')),
          },
        ],
        activeSheet: 'Daily schedule',
        selection: 'B2:I20',
      }),
    );
    const sc = scriptedClient([
      '```plan\nintent rewrite\nsurface excel\nscope selection\nstep add the proposed morning routine to the schedule\nconfidence high\n```',
    ]);
    ui = mountStack({ surface: 'excel', client: sc });
    await ui.flush();

    await typeAndSubmit('Okay add this to my schedule');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    expect(sc.queries[0]).toContain('Emit one ```plan block.');
    expect(sc.queries[0]).toContain('Okay add this to my schedule');
    expect(sc.queries[0]).not.toContain('COMMANDS — one per line');
    expect(ui.controller.getState().pendingCommandPlan?.plan.intent).toBe('rewrite');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });

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
    expect(card!.textContent).toContain('Here is what I plan to do');
    expect(card!.textContent).not.toContain('/rewrite');
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
    const threadText = ui.container.querySelector('.thread')?.textContent ?? '';
    expect(threadText).toContain('Confirmed the plan');
    expect(threadText).not.toContain('/execute');
    expect(threadText).not.toContain('<confirmed_plan>');
    expect(threadText).not.toContain('Treat the plan as approved intent');
    expect((sim as WordSimulator).snapshot().inserts.length).toBe(0); // still nothing applied — gated
  });

  it('a simple typed /rewrite still enters the planner first', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```plan\nintent rewrite\nsurface word\nscope selection\nstep make the selected text formal\nconfidence high\n```',
      ]),
    });
    await ui.flush();

    await typeAndSubmit('/rewrite make it formal');
    await ui.waitFor((s) => s.pendingCommandPlan !== undefined);

    expect(ui.container.querySelector('.command-plan')).not.toBeNull();
    expect(ui.controller.getState().pendingCommandPlan?.plan.intent).toBe('rewrite');
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
  });
});
