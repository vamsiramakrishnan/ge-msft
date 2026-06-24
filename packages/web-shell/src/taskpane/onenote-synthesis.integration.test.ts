// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import type { ProvenancePayload, SourceRef } from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import { scriptedClient, mountStack, type MountedStack } from '../test-harness/index.js';
import {
  installFakeOneNote,
  oneNoteSeed,
  oneNoteSection,
  oneNoteOutline,
  type OneNoteSimulator,
} from '../test-harness/fake-onenote.js';

/**
 * FULL-STACK OneNote **page-synthesis** integration tests. Each installs an in-memory OneNote host
 * (a seeded "Source review" page + an empty active section), then drives the REAL stack —
 * `selectBridge('onenote')` → real `AssistSession` (over a scripted fake client) → real
 * `PanelController` → rendered `<App/>` — through the page-synthesis path and reads the synthesized
 * page back FROM THE HOST.
 *
 * Seam under test: `@ge/bridge-onenote` ⇄ `@ge/runtime` ⇄ `@ge/web-shell`. Only the OneNote host
 * (globals) and the model stream (scripted SSE) are faked; everything between — the command grammar,
 * the plan-approval gate, `AssistSession.applyRequest`/`apply`, the `OneNoteBridge.actuate`
 * synthesis, the `ProvenanceStore` — is the real code.
 *
 * Two provenance surfaces are asserted, matching how OneNote actually carries provenance:
 *   1. **In the host page** — the appended page's outline HTML carries the synthesized claim as a
 *      block AND an inline `data-ge-cite` citation tag per grounding source (the per-claim `[source]`
 *      chips). Read back from the fake host's `addedPages` log.
 *   2. **In the session record** — the `ProvenanceStore` `ChangeRecord` the controller stamps from
 *      the turn's provenance carries the agent id, identity, content hash, and sources (who/what/why).
 */

let sim: OneNoteSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

/** A provenance payload with real attribution + grounding sources (the signed-in user). */
function provenanceWithSources(sources: SourceRef[]): ProvenancePayload {
  return {
    agentId: 'gemini-enterprise:research',
    identity: 'sim.user@acme',
    timestamp: '2026-06-24T00:00:00Z',
    sources,
    contentHash: 'sha256:onenote-synthesis',
    sessionId: 'sess_onenote' as never,
  };
}

describe('OneNote page-synthesis full-stack integration', () => {
  it('synthesizes a page: the plan-approved write lands a titled, body-bearing page in the host section', async () => {
    const section = oneNoteSection();
    sim = installFakeOneNote(
      oneNoteSeed({
        page: {
          id: 'pg-1',
          title: 'Source review',
          contents: [oneNoteOutline([{ type: 'RichText', text: 'Northwind MSA v3 is current.' }])],
        },
        section,
      }),
    );
    // The agent emits a single `page "<title>" "<body>"` command → OneNote `append-page`.
    ui = mountStack({
      surface: 'onenote',
      client: scriptedClient([
        '```cmd\npage "Risk synthesis" "The SLA sits below the customer standard."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('synthesize a risk page from the sources');
    });
    // A single append-page effect flows through the plan-approval gate (like Word's suggest).
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // The approval preview surfaces the append-page effect (kind + the synthesized body) the model
    // emitted, shown to the user before anything lands on the host.
    const planText = ui!.container.querySelector('.plan-approval')?.textContent ?? '';
    expect(planText).toContain('append-page');
    expect(planText).toContain('The SLA sits below the customer standard.');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // Read the synthesized page back FROM THE HOST: exactly one page was appended to the section,
    // titled from the command, carrying the claim as an outline block.
    const snap = sim.snapshot();
    expect(snap.addedPages).toHaveLength(1);
    expect(snap.addedPages[0]?.title).toBe('Risk synthesis');
    const html = snap.addedPages[0]?.outlineHtml ?? '';
    expect(html).toContain('The SLA sits below the customer standard.');
    // The outline is a real HTML paragraph block (not a bare string dump).
    expect(html).toContain('<p>');
    // The outline is positioned at the bridge's standard top-left offset on the new page.
    expect(snap.addedPages[0]?.outlineLeft).toBe(40);
    expect(snap.addedPages[0]?.outlineTop).toBe(40);
  });

  it('treats the synthesized body as UNTRUSTED data — HTML in the agent text is escaped on the host page, never injected', async () => {
    const section = oneNoteSection();
    sim = installFakeOneNote(
      oneNoteSeed({
        page: { id: 'pg-x', title: 'Source review', contents: [] },
        section,
      }),
    );
    // The model-shaped body is untrusted (its text is influenced by document content); a markup
    // payload must be screened to data, not rendered as live HTML on the page.
    ui = mountStack({
      surface: 'onenote',
      client: scriptedClient([
        '```cmd\npage "Injected" "<img src=x onerror=alert(1)> review"\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('synthesize a page');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const html = sim.snapshot().addedPages[0]?.outlineHtml ?? '';
    // The injected tag is escaped (rendered as data), and no live <img> tag reaches the host page.
    expect(html).toContain('&lt;img src=x onerror=alert(1)&gt;');
    expect(html).not.toContain('<img');
  });

  it('a citation-tagged synthesis lands in the host AND its provenance (agent/identity/hash/sources) is recorded in-session', async () => {
    const section = oneNoteSection();
    sim = installFakeOneNote(
      oneNoteSeed({
        page: {
          id: 'pg-2',
          title: 'Source review',
          contents: [
            oneNoteOutline([{ type: 'RichText', text: 'ISO 27001 valid through Nov 2026.' }]),
          ],
        },
        section,
      }),
    );

    const sources: SourceRef[] = [{ title: 'Vendor Risk Policy v4', locator: '§3.2' }];
    // A grounded turn that emits the source as a citation AND stamps a provenance payload carrying
    // who (identity), what (agentId), the content hash, and the grounding sources.
    ui = mountStack({
      surface: 'onenote',
      client: scriptedClient(
        [{ text: 'The vendor SLA is below standard.', citations: sources }],
        provenanceWithSources(sources),
      ),
    });
    await ui!.flush();

    // 1) Run a grounded ask so the turn's provenance becomes resident on the controller.
    await ui!.act(() => void ui!.controller.send('how does the SLA compare to standard?'));
    await ui!.waitFor((s) => !s.busy);
    await ui!.flush();

    // 2) Stage + apply the append-page synthesis carrying the grounding sources (the citation tags).
    let proposalId!: ReturnType<MountedStack['controller']['propose']>['changeId'];
    await ui!.act(() => {
      const proposal = ui!.controller.propose(
        'append-page',
        {
          text: 'The vendor SLA is below the customer standard and needs escalation.',
          sources,
          target: { matchText: 'SLA risk synthesis' },
        },
        'Synthesize SLA risk page',
      );
      proposalId = proposal.changeId;
    });
    await ui!.act(() => void ui!.controller.applyProposal(proposalId));
    await ui!.waitFor((s) => s.changes.length > 0);
    await ui!.flush();

    // Provenance surface #1 — IN THE HOST: read the appended page back. Its outline carries the
    // synthesized claim AND an inline citation tag derived from the grounding source.
    const snap = sim.snapshot();
    expect(snap.addedPages).toHaveLength(1);
    expect(snap.addedPages[0]?.title).toBe('SLA risk synthesis');
    const html = snap.addedPages[0]?.outlineHtml ?? '';
    expect(html).toContain('The vendor SLA is below the customer standard and needs escalation.');
    expect(html).toContain('data-ge-cite="1"');
    expect(html).toContain('[Vendor Risk Policy v4 · §3.2]');

    // Provenance surface #2 — IN-SESSION RECORD: the ChangeRecord carries who/what/why for the write.
    const changes = ui!.controller.getState().changes;
    const record = changes.find((c) => c.changeId === proposalId);
    expect(record).toBeDefined();
    expect(record?.ok).toBe(true);
    expect(record?.kind).toBe('append-page');
    expect(record?.provenance?.agentId).toBe('gemini-enterprise:research');
    expect(record?.provenance?.identity).toBe('sim.user@acme');
    expect(record?.provenance?.contentHash).toBe('sha256:onenote-synthesis');
    expect(record?.provenance?.sources).toEqual(sources);
    // The host write location echoes the appended page id (traceable back to the host).
    expect(record?.location).toBe(`page:${snap.addedPages[0]?.id}`);
  });

  it('degrades cleanly when there is NO active section: nothing is written to the host and the failure is recorded', async () => {
    // No active section → the OneNote bridge degrades the append-page rather than throwing.
    sim = installFakeOneNote(
      oneNoteSeed({
        page: { id: 'pg-3', title: 'Source review', contents: [] },
        section: null,
      }),
    );

    ui = mountStack({ surface: 'onenote', client: scriptedClient([]) });
    await ui!.flush();

    let proposalId!: ReturnType<MountedStack['controller']['propose']>['changeId'];
    await ui!.act(() => {
      const proposal = ui!.controller.propose(
        'append-page',
        { text: 'An ungrounded synthesis that cannot land.', target: { matchText: 'Orphan' } },
        'Synthesize orphan page',
      );
      proposalId = proposal.changeId;
    });
    await ui!.act(() => void ui!.controller.applyProposal(proposalId));
    await ui!.waitFor((s) => s.changes.length > 0);
    await ui!.flush();

    // Nothing reached the host (no section to append into), and the degraded failure is recorded.
    expect(sim.snapshot().addedPages).toHaveLength(0);
    const record = ui!.controller.getState().changes.find((c) => c.changeId === proposalId);
    expect(record?.ok).toBe(false);
    expect(record?.degraded).toBe(true);
    expect(record?.error?.code).toBe('no_section');
  });

  it('passes the real change id through the seam: the controller-minted id is echoed back from the bridge result', async () => {
    const section = oneNoteSection();
    sim = installFakeOneNote(
      oneNoteSeed({ page: { id: 'pg-4', title: 'Source review', contents: [] }, section }),
    );
    ui = mountStack({ surface: 'onenote', client: scriptedClient([]) });
    await ui!.flush();

    let proposalId!: ReturnType<MountedStack['controller']['propose']>['changeId'];
    await ui!.act(() => {
      const proposal = ui!.controller.propose(
        'append-page',
        { text: 'A grounded claim.', target: { matchText: 'Echo' } },
        'Echo page',
      );
      proposalId = proposal.changeId;
    });
    await ui!.act(() => void ui!.controller.applyProposal(proposalId));
    await ui!.waitFor((s) => s.changes.length > 0);
    await ui!.flush();

    const record = ui!.controller.getState().changes.find((c) => c.changeId === proposalId);
    expect(record?.ok).toBe(true);
    // The change id the controller minted survives the full round-trip through the bridge.
    expect(record?.changeId).toBe(proposalId);
    expect(proposalId).toBe(asChangeId(proposalId));
  });
});
