// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { quickActionsForSurface } from '@ge/contracts';
import { quickActionSeed } from './components/quick-action-seed.js';
import {
  installFakeWord,
  scriptedClient,
  mountStack,
  type WordSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * FULL-STACK interplay for the command surface. Each test installs an in-memory Word host, then
 * drives the REAL stack — `selectBridge('word')` → real `AssistSession` (scripted client) → real
 * `PanelController` → rendered `<App/>` (so `QuickActionBar` + `Composer` are the real components) —
 * and asserts that a quick-action chip / a `/verb` submit routes to the correct controller path:
 *   - read-only `chat` actions and plain questions → `send` (a grounded turn, no gate);
 *   - `write`/`annotation` actions and actuating `/verbs` → `runCommands` → the fail-closed plan gate.
 * No new actuation path is introduced by the chips/palette; they only seed the existing routes.
 */

let sim: WordSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

function chip(actionId: string): HTMLButtonElement {
  const el = ui!.container.querySelector<HTMLButtonElement>(
    `button.quick-action[data-action-id="${actionId}"]`,
  );
  if (!el) throw new Error(`no quick-action chip with data-action-id "${actionId}"`);
  return el;
}

async function typeAndSubmit(text: string): Promise<void> {
  const input = ui!.container.querySelector<HTMLInputElement>('input#ask')!;
  await ui!.act(() => {
    const setter = Object.getOwnPropertyDescriptor(
      window.HTMLInputElement.prototype,
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
    await ui.act(() => chip('summarize-this').click());
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
    await ui.act(() => chip('tighten').click());
    await ui.waitFor((s) => s.pendingPlan !== undefined);

    // The write surfaced as a plan-approval card (the gate) rather than auto-applying.
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
    expect(sim.snapshot().inserts.length).toBe(0); // nothing applied before approval
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
});

describe('command surface — planner-confirm for complex free-text (full-stack, §F)', () => {
  it('a complex /rewrite proposes a plan, then runs the executor on confirm', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        // 1) the PLANNER turn → a ```plan block (the confirmable intention)
        '```plan\nintent rewrite\nsurface word\nstep rewrite the SLA to 99.9% as a tracked change\nexclude the indemnity clause\n```',
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
    expect(ui.controller.getState().pendingPlan).toBeUndefined();
    expect(sim.snapshot().inserts.length).toBe(0);

    // Confirm → the executor runs and stages ITS OWN effect-level gate.
    await ui.act(() =>
      ui!.container
        .querySelector<HTMLButtonElement>('[data-testid="command-plan-confirm"]')!
        .click(),
    );
    await ui.waitFor((s) => s.pendingPlan !== undefined);
    expect(ui.controller.getState().pendingCommandPlan).toBeUndefined();
    expect(ui.container.querySelector('.plan-approval')).not.toBeNull();
    expect(sim.snapshot().inserts.length).toBe(0); // still nothing applied — gated
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
