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
 * FULL-STACK interplay: bridge-word + runtime + web-shell.
 *
 * Track: stale-anchor-degrade. The existing `word-integration.test.ts` covers the case where the
 * content anchor was ABSENT FROM THE START (never resolvable). This file covers the harder, distinct
 * case the anchoring contract (CLAUDE.md: "re-resolve at apply-time") is really about: a `body.search`
 * anchor that RESOLVES at dry-run/preview time but goes STALE *between* the plan being staged and the
 * user approving it — a coauthor (or the user) mutates the document in the approval window.
 *
 * We drive the REAL stack — `selectBridge('word')` → real `AssistSession` (scripted client) → real
 * `PanelController` → rendered `<App/>` — and mock only the two outer boundaries: the Word HOST (the
 * in-memory `installFakeWord` simulator) and the model NETWORK (the scripted SSE client). Everything
 * between is real: real plan dry-run, real plan-approval gate, real `WordBridge.applyTrackedChange`,
 * real apply-time `body.search` re-resolution, real `ProvenanceStore`.
 *
 * The seam assertion: because the bridge re-resolves the anchor against the LIVE body inside the
 * apply-time `Word.run` batch (not against a range id captured at preview), mutating the fake
 * document after the plan is staged makes the previously-live anchor drift. The write must DEGRADE to
 * a panel item (`anchor_drift` surfaced in the run transcript), insert NOTHING into the body, and
 * leave NO durable provenance — rather than render a broken annotation on the wrong range.
 */

let sim: WordSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

const ANCHOR = 'The SLA is 99.5%.';
const SUGGESTION = 'The SLA is ~99.5% (source needed).';

/** A document that DOES contain the anchor — so it resolves at preview, before we mutate it. */
function liveAnchorSeed(): ReturnType<typeof wordSeed> {
  return wordSeed({
    selectionText: ANCHOR,
    paragraphs: [
      { text: 'Master Services Agreement', styleBuiltIn: 'Heading1' },
      { text: '1. Service Levels', styleBuiltIn: 'Heading2' },
      {
        text: `${ANCHOR} Downtime beyond this entitles the customer to service credits.`,
        styleBuiltIn: 'Normal',
      },
    ],
  });
}

describe('Word stale-anchor degrade (full-stack interplay)', () => {
  it('an anchor that resolves at preview but goes STALE before apply degrades to a panel item', async () => {
    sim = installFakeWord(liveAnchorSeed());
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        `\`\`\`cmd\nsuggest "${ANCHOR}" => "${SUGGESTION}"\n\`\`\``,
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag the unsourced SLA claim');
    });

    // ── PREVIEW: the plan is staged. The anchor is LIVE right now. ─────────────
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The plan-approval card renders the verbatim suggest command targeting the live anchor text.
    const cmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(cmd).toContain('suggest');
    expect(cmd).toContain(ANCHOR);

    // Cross-boundary proof the anchor is genuinely resolvable at preview-time: drive the REAL bridge's
    // content-anchored `body.search` (the same re-resolution the apply path uses) against the live
    // host — it finds the anchor NOW, so this is a true stale-after-preview drift, not absent-from-start.
    // (`searchDocument` is an optional `DocBridge` capability; the Word bridge implements it.)
    expect(ui!.bridge.searchDocument).toBeDefined();
    const hitsAtPreview = await ui!.bridge.searchDocument!(ANCHOR);
    expect(hitsAtPreview.length).toBeGreaterThan(0);

    // ── DRIFT: a coauthor rewrites the anchored paragraph in the approval window. ──
    // Mutate the SAME mutable seed the bridge reads, so the apply-time re-search misses.
    const anchored = sim.seed.paragraphs.find((p) => p.text.includes(ANCHOR));
    if (!anchored) throw new Error('expected the anchored paragraph to be present at preview-time');
    anchored.text = 'This revision removes the service-level commitment entirely.';

    // ── APPLY: approve the plan. The bridge re-resolves at apply-time and finds nothing. ──
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const snap = sim.snapshot();
    // The drift DEGRADED: nothing was inserted as a tracked change, and the suggestion text never
    // landed in the body — no broken annotation on a stale or wrong range.
    expect(snap.inserts.length).toBe(0);
    expect(snap.bodyText).not.toContain(SUGGESTION);
    expect(snap.bodyText).not.toContain('source needed');
    // The coauthor's edit is still there (we did not clobber it).
    expect(snap.bodyText).toContain('removes the service-level commitment');

    // The run transcript surfaces the target-conflict / regenerate signal: the degraded write
    // narrates the `anchor_drift` error code rather than throwing or silently no-op-ing.
    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).toContain('anchor_drift');

    // A degraded write records NO durable provenance custom XML part (provenance lands only on an
    // applied, non-degraded reversible write).
    expect(sim.office.customXmlParts.length).toBe(0);
  });

  it('the SAME suggestion applies cleanly when the anchor stays live through approval (control)', async () => {
    // Control: identical plan, but no drift in the approval window → the tracked change DOES land.
    // This proves the degradation above is caused by the stale anchor, not by the plan path itself.
    sim = installFakeWord(liveAnchorSeed());
    ui = mountStack({
      surface: 'word',
      client: scriptedClient([
        `\`\`\`cmd\nsuggest "${ANCHOR}" => "${SUGGESTION}"\n\`\`\``,
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag the unsourced SLA claim');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // No mutation in the approval window — the anchor stays live.
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const snap = sim.snapshot();
    expect(snap.inserts.length).toBe(1);
    expect(snap.inserts[0]?.anchor).toBe(ANCHOR);
    expect(snap.bodyText).toContain(SUGGESTION);
    expect(snap.changeTrackingMode).toBe('TrackAll');
    // The clean apply DID record durable provenance (the stale case above recorded none).
    expect(sim.office.customXmlParts.length).toBeGreaterThan(0);

    const steps = ui!.container.querySelector('.run-steps')?.textContent ?? '';
    expect(steps).not.toContain('anchor_drift');
  });
});
