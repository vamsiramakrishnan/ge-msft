// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import {
  installFakeWord,
  wordSeed,
  scriptedClient,
  mountStack,
  type WordSimulator,
  type MountedStack,
  openPlanEffects,
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

    await ui!.showSources();
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
    openPlanEffects(ui!.container);
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
    // Only the changed word is replaced (tracked), not the whole sentence (UAT #17).
    expect(snap.inserts.length).toBe(1);
    expect(snap.inserts[0]).toMatchObject({ anchor: '99.5%.', text: '~99.5% (source needed).' });
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

  it('/insert-text with bold, font and size sets them on the inserted text (UAT #20)', async () => {
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\n/insert-text text="AI COE" bold=true font=Arial size=16\n```',
        '```cmd\ndone "Inserted."\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('tulis AI COE bold, arial, 16');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    const insert = sim.snapshot().inserts.at(-1);
    expect(insert?.text).toBe('AI COE');
    expect(insert?.font).toEqual({ bold: true, name: 'Arial', size: 16 });
  });

  it('turns paragraphs into a table: one tracked edit replaces them (WS4b, Argon mx-wd-table-14)', async () => {
    sim = installFakeWord(
      wordSeed({
        paragraphs: [
          { text: 'Action items', styleBuiltIn: 'Heading2' },
          { text: 'Budi - send the test results - 21 Oct', styleBuiltIn: 'Normal' },
          { text: 'Rina - review the SLA - 24 Oct', styleBuiltIn: 'Normal' },
          { text: 'Next steps follow.', styleBuiltIn: 'Normal' },
        ],
      }),
    );
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\n/insert-table match="Budi - send" replaceThrough="Rina - review" rows="Owner\\tAction\\tDue\\nBudi\\tSend the test results\\t21 Oct\\nRina\\tReview the SLA\\t24 Oct" headers=true\n```',
        '```cmd\ndone "Converted."\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('turn the action items into a table: owner, action, due');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    const doc = sim.snapshot();
    expect(doc.tables?.at(-1)?.[0]).toEqual(['Owner', 'Action', 'Due']);
    const texts = doc.paragraphs.map((p) => p.text);
    expect(texts).not.toContain('Budi - send the test results - 21 Oct');
    expect(texts).not.toContain('Rina - review the SLA - 24 Oct');
    expect(texts[0]).toBe('Action items');
    expect(texts.at(-1)).toBe('Next steps follow.');
    // The source paragraphs went out as tracked deletions (the user's own mode is restored after).
    const removed = doc.inserts.filter((i) => i.text === '');
    expect(removed.map((i) => i.anchor)).toEqual([
      'Budi - send the test results - 21 Oct',
      'Rina - review the SLA - 24 Oct',
    ]);
    expect(removed.every((i) => i.tracked)).toBe(true);
  });

  it('refuses to replace more than 50 paragraphs with a table (WS4b)', async () => {
    const paragraphs = Array.from({ length: 60 }, (_, i) => ({
      text: `Line ${i + 1}`,
      styleBuiltIn: 'Normal',
    }));
    sim = installFakeWord(wordSeed({ paragraphs }));
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\n/insert-table match="Line 1" replaceThrough="Line 60" rows="A\\tB\\n1\\t2"\n```',
        '```cmd\ndone "Too long."\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('table everything');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    expect(sim.snapshot().paragraphs).toHaveLength(60);
  });

  it('find-replace honours matchWholeWord: "monthly" stays, each "month" changes once (WS5 T2b)', async () => {
    sim = installFakeWord(
      wordSeed({
        paragraphs: [
          { text: 'Fees are billed each month and reported monthly.', styleBuiltIn: 'Normal' },
          { text: 'The month ends on the last day of the month.', styleBuiltIn: 'Normal' },
        ],
      }),
    );
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        '```cmd\n/find-replace find="month" replace="calendar month" matchWholeWord=true\n```',
        '```cmd\ndone "Replaced."\n```',
      ]),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('replace month with calendar month');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    expect(sim.snapshot().paragraphs.map((p) => p.text)).toEqual([
      'Fees are billed each calendar month and reported monthly.',
      'The calendar month ends on the last day of the calendar month.',
    ]);
  });

  describe('replacing paragraphs with a table fails safe (security review H2, M1, M4)', () => {
    const doc = (
      extra: Partial<Record<'inTable' | 'sectionBreak' | 'lockedControl', boolean>> = {},
    ) =>
      wordSeed({
        paragraphs: [
          { text: 'Intro', styleBuiltIn: 'Normal' },
          { text: 'Budi - send the test results', styleBuiltIn: 'Normal', ...extra },
          { text: 'Rina - review the test results', styleBuiltIn: 'Normal' },
          { text: 'Outro', styleBuiltIn: 'Normal' },
        ],
      });
    const convert = async (cmd: string, task = 'make a table') => {
      ui = mountStack({
        surface: 'word',
        client: scriptedClient([`\`\`\`cmd\n${cmd}\n\`\`\``, '```cmd\ndone "ok"\n```']),
      });
      await ui!.flush();
      let run!: Promise<void>;
      await ui!.act(() => {
        run = ui!.controller.runCommands(task);
      });
      await ui!.waitFor((s) => s.pendingPlan !== undefined);
      await ui!.act(() => ui!.controller.approvePlan());
      await ui!.waitFor((s) => !s.busy);
      await run;
      return sim!.snapshot().paragraphs.map((p) => p.text);
    };
    const table = 'rows="A\\tB\\n1\\t2"';

    it('refuses on a host that cannot force Track Changes (WordApi < 1.6)', async () => {
      sim = installFakeWord(doc(), { WordApi: 5 });
      const texts = await convert(`/insert-table match="Budi" replaceThrough="Rina" ${table}`);
      expect(texts).toHaveLength(4);
    });

    it('ends at the nearest paragraph: end text in the start paragraph replaces only that one', async () => {
      sim = installFakeWord(doc());
      const texts = await convert(
        `/insert-table match="Budi - send" replaceThrough="test results" ${table}`,
      );
      expect(texts).toEqual(['Intro', 'A\tB', '1\t2', 'Rina - review the test results', 'Outro']);
    });

    it.each(['inTable', 'sectionBreak', 'lockedControl'] as const)(
      'refuses a run that crosses %s, changing nothing',
      async (flag) => {
        sim = installFakeWord(doc({ [flag]: true }));
        const texts = await convert(`/insert-table match="Budi" replaceThrough="Rina" ${table}`);
        expect(texts).toHaveLength(4);
        expect(sim.snapshot().tables ?? []).toEqual([]);
      },
    );
  });
});
