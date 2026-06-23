// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  installFakeWord,
  wordSeed,
  scriptedClient,
  mountStack,
  type WordSimulator,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * FULL-STACK Word integration tests. Each installs an in-memory Word host (seeded document), then
 * drives the REAL stack — `selectBridge('word')` → real `AssistSession` (scripted fake client) →
 * real `PanelController` → rendered `<App/>` — and asserts the UI shows the SIMULATED document and a
 * content-anchored tracked change resolves against the seeded body (and degrades on a stale anchor).
 *
 * Word's tracked change is a single effect, so it flows through the plan-approval gate (the
 * controller wires `approvePlan`, which takes precedence over the per-write gate). The card renders
 * the verbatim `suggest "old" => "new"` command line the model emitted.
 */

let sim: WordSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

describe('Word full-stack integration', () => {
  it('refreshContext surfaces the seeded selection + whole document as context chips', async () => {
    sim = installFakeWord();
    ui = mountStack({ surface: 'word', client: scriptedClient([]) });
    await ui!.flush();
    await ui!.act(() => void ui!.controller.refreshContext());
    await ui!.flush();

    const text = ui!.container.textContent ?? '';
    // The seeded selection text previews on its chip; the whole-document chip is always present.
    expect(text).toContain('Selection');
    expect(text).toContain('Whole document');
  });

  it('a suggest tracked change renders the verbatim command and resolves against the seeded body', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (source needed)."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag the unsourced SLA claim');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The approval card shows the verbatim suggest command (old => new), shown to the user exactly.
    const cmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(cmd).toContain('suggest');
    expect(cmd).toContain('The SLA is 99.5%.');
    expect(cmd).toContain('source needed');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The tracked change resolved against the seeded paragraph and rewrote it in place.
    const snap = sim.snapshot();
    expect(snap.inserts.length).toBe(1);
    expect(snap.inserts[0]?.anchor).toBe('The SLA is 99.5%.');
    expect(snap.bodyText).toContain('The SLA is ~99.5% (source needed).');
    // Tracked changes were turned on before the edit.
    expect(snap.changeTrackingMode).toBe('TrackAll');
    // Durable provenance landed as a custom XML part (Word's durable metadata).
    expect(sim.office.customXmlParts.length).toBeGreaterThan(0);
    expect(sim.office.customXmlParts[0]).toContain('sim.user@acme');
  });

  it('a suggest on an ABSENT anchor degrades to a panel item (no broken annotation)', async () => {
    // Seed a document WITHOUT the anchor text so the apply-time re-resolution finds no hit.
    sim = installFakeWord(
      wordSeed({
        paragraphs: [
          { text: 'Master Services Agreement', styleBuiltIn: 'Heading1' },
          {
            text: 'There is no service-level commitment in this revision.',
            styleBuiltIn: 'Normal',
          },
        ],
      }),
    );
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (source needed)."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag the SLA claim');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The drift degraded: nothing was inserted, and the body is untouched.
    const snap = sim.snapshot();
    expect(snap.inserts.length).toBe(0);
    expect(snap.bodyText).not.toContain('source needed');
    // The run-step transcript narrates the drifted (degraded) write rather than throwing: the
    // controller surfaces the failure's error code for a non-ok result.
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).toContain('anchor_drift');
  });

  it('rejecting the plan writes nothing (per-plan gating holds for Word too)', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is 99.9%."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('bump the SLA');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.rejectPlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const snap = sim.snapshot();
    expect(snap.inserts.length).toBe(0);
    expect(snap.bodyText).toContain('The SLA is 99.5%.'); // unchanged
    expect(sim.office.customXmlParts.length).toBe(0); // no provenance for a rejected write
  });
});
