// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  installFakePowerPoint,
  scriptedClient,
  mountStack,
  type PowerPointSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * FULL-STACK PowerPoint integration tests (lighter than Excel/Word). Installs an in-memory deck,
 * drives the REAL stack — `selectBridge('powerpoint')` → real `AssistSession` (scripted fake client)
 * → real `PanelController` → `<App/>` — and asserts the UI lists the seeded slide and an approved
 * `slide` insert composes a new slide into the fake deck.
 */

let sim: PowerPointSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

describe('PowerPoint full-stack integration', () => {
  it('refreshContext surfaces the selected slide + whole deck as context chips', async () => {
    sim = installFakePowerPoint();
    ui = mountStack({ surface: 'powerpoint', client: scriptedClient([]) });
    await ui!.flush();
    await ui!.act(() => void ui!.controller.refreshContext());
    await ui!.flush();

    const text = ui!.container.textContent ?? '';
    // The default deck selects slide index 1 → "Slide 2"; the whole-deck chip is always present.
    expect(text).toContain('Slide 2');
    expect(text).toContain('Whole deck');
  });

  it('an approved slide insert composes a new slide into the fake deck', async () => {
    sim = installFakePowerPoint();
    const before = sim.snapshot().slides.length;
    ui = mountStack({
      surface: 'powerpoint',
      client: scriptedClient([
        '```cmd\nslide "Q4 Outlook" "Revenue guidance raised" "Hiring freeze lifted"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('add a closing slide');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The plan card stages the insert-slide effect for one approval. (The verbatim command renderer
    // has no dedicated insert-slide line, so it labels the effect by kind — the slide title lives in
    // params.slide, which the host write consumes; we assert the title landed on the deck below.)
    const plan = ui!.container.querySelector('.plan-approval');
    expect(plan).not.toBeNull();
    expect(plan?.querySelector('.pin')?.textContent).toContain('insert-slide');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The fake deck grew by one slide whose shapes carry the composed title + bullets.
    const snap = sim.snapshot();
    expect(snap.slides.length).toBe(before + 1);
    const added = snap.slides[snap.slides.length - 1];
    expect(added?.shapeTexts[0]).toBe('Q4 Outlook');
    expect(added?.shapeTexts[1]).toContain('Revenue guidance raised');
    expect(added?.shapeTexts[1]).toContain('Hiring freeze lifted');
  });

  it('rejecting the plan leaves the deck unchanged', async () => {
    sim = installFakePowerPoint();
    const before = sim.snapshot().slides.length;
    ui = mountStack({
      surface: 'powerpoint',
      client: scriptedClient(['```cmd\nslide "Unwanted" "nope"\n```', '```cmd\ndone\n```']),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('add a slide');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.rejectPlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    expect(sim.snapshot().slides.length).toBe(before);
  });
});
