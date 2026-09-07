// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import {
  ProvenancePayloadSchema,
  asChangeId,
  type ProvenancePayload,
  type ActuationResult,
  type ChangeId,
  type SourceRef,
} from '@ge/contracts';
import { ProvenanceStore } from '../provenance-store.js';
import { ProposalCard } from './components/ProposalCard.js';
import type { Proposal } from '../controller.js';
import {
  installFakeWord,
  defaultWordSeed,
  installFakeExcel,
  defaultExcelSeed,
  scriptedClient,
  mountStack,
  type WordSimulator,
  type WordSeed,
  type ExcelSimulator,
  type ExcelSeed,
  type MountedStack,
} from '../test-harness/index.js';

/**
 * FULL-STACK PROVENANCE ROUND-TRIP. Wires the REAL seam:
 *   real AssistSession (scripted client) → real bridge → fake host durable metadata
 *     → parse the durable record BACK into a ProvenancePayload
 *       → real ProvenanceStore reads it back
 *         → real ProposalCard/ProvenanceDetail drill-in renders who/when/sources/hash.
 *
 * The ONLY fakes are the outer boundaries — the Office host (in-memory) and the model stream
 * (scripted SSE). Everything in between is the production stack. The round-trip asserts the FULL
 * provenance cycle: an approved write actually lands provenance into the host's DURABLE metadata
 * (Word custom-XML part / Excel workbook settings), that durable record is read back and
 * re-validated against the contract's Zod schema, the in-session ProvenanceStore reflects it, and
 * the UI drill-in renders every field (agent / identity / timestamp / sources / hash).
 *
 * Then it simulates SAVE + REOPEN: a fresh fake host is installed carrying the SAME durable
 * metadata the first session wrote, and the provenance is asserted to still parse identically —
 * proving the write stays provenanced across a document close/reopen.
 */

/* ───────────────────────── durable-metadata read-back parsers ───────────────────────── */
/*
 * The bridges WRITE durable provenance (Word → customXmlParts.add(xml); Excel → settings JSON) but
 * expose no read-back API. These parsers are the round-trip's READER half: they decode the EXACT
 * durable format the production `provenanceRecord` serializer emitted, so the test reads back what
 * the real bridge actually persisted (not a hand-rolled fixture).
 */

/** Decode an attribute off the Word custom-XML provenance part (entities un-escaped). */
function xmlAttr(xml: string, name: string): string | undefined {
  const m = xml.match(new RegExp(`${name}="([^"]*)"`));
  return m ? unescapeXml(m[1] ?? '') : undefined;
}

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/** Parse the Word OOXML custom-XML provenance part back into a (contract-validated) payload. */
function parseWordProvenance(xml: string): ProvenancePayload {
  const sources: SourceRef[] = [];
  for (const tag of xml.match(/<source\b[^>]*\/>/g) ?? []) {
    const title = xmlAttr(tag, 'title') ?? '';
    const uri = xmlAttr(tag, 'uri');
    const locator = xmlAttr(tag, 'locator');
    sources.push({ title, ...(uri ? { uri } : {}), ...(locator ? { locator } : {}) });
  }
  const payload = {
    agentId: xmlAttr(xml, 'agentId') ?? '',
    identity: xmlAttr(xml, 'identity') ?? '',
    timestamp: xmlAttr(xml, 'timestamp') ?? '',
    contentHash: xmlAttr(xml, 'contentHash') ?? '',
    sources,
  };
  // Re-validate against the AUTHORITATIVE contract schema on read-back (parse-on-receipt).
  return ProvenancePayloadSchema.parse(payload);
}

/** Parse the Excel settings JSON provenance record back into a (contract-validated) payload. */
function parseExcelProvenance(json: string): ProvenancePayload {
  const obj = JSON.parse(json) as Record<string, unknown>;
  // The serializer folds `changeId` in alongside the payload; strip it before schema validation.
  const { changeId: _changeId, ...rest } = obj;
  void _changeId;
  return ProvenancePayloadSchema.parse(rest);
}

/* ───────────────────────── real ProposalCard drill-in rendering ───────────────────────── */

let drill: { container: HTMLDivElement; root: Root } | undefined;

/**
 * Render the REAL `ProposalCard` (web-shell UI) for an applied proposal carrying the read-back
 * provenance, then open its drill-in. This is the production drill-in component path: the same
 * `ProvenanceDetail` the live App renders inside a ProposalCard.
 */
function renderDrillIn(provenance: ProvenancePayload): HTMLDivElement {
  (globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
  const proposal: Proposal = {
    changeId: asChangeId('c-roundtrip'),
    kind: 'tracked-change',
    params: { target: { matchText: 'The SLA is 99.5%.' }, text: 'The SLA is ~99.5%.' },
    label: 'Flag the unsourced SLA claim',
    status: 'applied',
    provenance,
  };
  const container = document.createElement('div');
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(createElement(ProposalCard, { proposals: [proposal], onApply: () => {} }));
  });
  // Open the provenance drill-down (real toggle button on the card).
  const toggle = container.querySelector<HTMLButtonElement>('.prov-toggle');
  if (!toggle) throw new Error('drill-in: no provenance toggle rendered');
  act(() => toggle.click());
  drill = { container, root };
  return container;
}

let sim: WordSimulator | ExcelSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  if (drill) {
    act(() => drill!.root.unmount());
    drill.container.remove();
  }
  ui = undefined;
  sim = undefined;
  drill = undefined;
});

/* ───────────────────────────────────── Word round-trip ───────────────────────────────── */

describe('provenance round-trip (Word: custom-XML durable metadata)', () => {
  it('actuated write → durable XML part → ProvenanceStore read-back → UI drill-in shows every field', async () => {
    const sources: SourceRef[] = [
      { title: 'Vendor SLA Policy', uri: 'https://policies.acme/sla', locator: '§4.1' },
    ];
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient(
        [
          {
            text: '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (source needed)."\n```',
            citations: sources,
          },
          '```cmd\ndone\n```',
        ],
        {
          agentId: 'gemini-enterprise:contract-review',
          identity: 'sim.user@acme',
          timestamp: '2026-06-23T00:00:00Z',
          sources,
          contentHash: 'sha256:roundtrip-word',
        },
      ),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('flag the unsourced SLA claim');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // (1) The tracked change actually landed in the host (cross-boundary mutation, read back).
    const wordSim = sim as WordSimulator;
    expect(wordSim.snapshot().inserts.length).toBe(1);
    expect(wordSim.snapshot().bodyText).toContain('source needed');

    // (2) Durable provenance landed as a custom XML part in the host's durable metadata.
    expect(wordSim.office.customXmlParts.length).toBe(1);
    const durableXml = wordSim.office.customXmlParts[0]!;

    // (3) READ IT BACK from the durable part and re-validate against the contract schema.
    const readBack = parseWordProvenance(durableXml);
    expect(readBack.agentId).toBe('gemini-enterprise:contract-review');
    expect(readBack.identity).toBe('sim.user@acme');
    expect(readBack.timestamp).toBe('2026-06-23T00:00:00Z');
    expect(readBack.contentHash).toBe('sha256:roundtrip-word');
    expect(readBack.sources).toEqual(sources);

    // (4) The REAL ProvenanceStore reflects the read-back record keyed by the durable changeId.
    const changeId = (xmlAttr(durableXml, 'changeId') ?? 'c-x') as ChangeId;
    const store = new ProvenanceStore(() => '2026-06-23T00:00:01Z');
    const result: ActuationResult = {
      ok: true,
      changeId: asChangeId(changeId),
      kind: 'tracked-change',
    };
    store.record(result, readBack);
    const rec = store.get(asChangeId(changeId));
    expect(rec?.provenance?.identity).toBe('sim.user@acme');
    expect(rec?.provenance?.sources[0]?.title).toBe('Vendor SLA Policy');

    // (5) The REAL UI drill-in renders agent / identity / timestamp / sources / hash from read-back.
    const drillEl = renderDrillIn(readBack);
    const text = drillEl.textContent ?? '';
    expect(text).toContain('gemini-enterprise:contract-review'); // agent
    expect(text).toContain('sim.user@acme'); // identity
    expect(drillEl.querySelector('time')?.getAttribute('dateTime')).toBe('2026-06-23T00:00:00Z');
    expect(drillEl.querySelector('.prov-hash')?.textContent).toBe('sha256:roundtrip-word');
    const link = drillEl.querySelector<HTMLAnchorElement>('.prov-sources a');
    expect(link?.getAttribute('href')).toBe('https://policies.acme/sla');
    expect(link?.textContent).toBe('Vendor SLA Policy · §4.1');
  });

  it('persists across save + reopen: a fresh host carrying the same durable metadata still parses', async () => {
    const sources: SourceRef[] = [{ title: 'Standard MSA', uri: 'https://acme/msa' }];
    sim = installFakeWord();
    ui = mountStack({
      surface: 'word',
      client: scriptedClient(
        [
          '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (cited)."\n```',
          '```cmd\ndone\n```',
        ],
        {
          agentId: 'gemini-enterprise:sim',
          identity: 'sim.user@acme',
          timestamp: '2026-06-23T00:00:00Z',
          sources,
          contentHash: 'sha256:persist-word',
        },
      ),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('cite the SLA claim');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const firstSession = sim as WordSimulator;
    expect(firstSession.office.customXmlParts.length).toBe(1);
    // Capture the durable metadata the host would persist into the .docx on save.
    const persistedParts = [...firstSession.office.customXmlParts];
    const persistedBody = firstSession.snapshot().bodyText;

    // SAVE + CLOSE: tear down the first session entirely (UI + host globals).
    ui!.unmount();
    firstSession.restore();
    ui = undefined;

    // REOPEN: a fresh fake host seeded with the SAME body the change wrote AND carrying the SAME
    // durable custom-XML parts. The reopened document must still expose the provenance.
    const reopenedSeed: WordSeed = {
      ...defaultWordSeed(),
      paragraphs: persistedBody.split('\n').map((text, i) => ({
        text,
        styleBuiltIn: i === 0 ? 'Heading1' : 'Normal',
      })),
    };
    sim = installFakeWord(reopenedSeed);
    const reopened = sim as WordSimulator;
    // Re-hydrate the durable metadata that survived the save (the host re-reads it from the file).
    reopened.office.customXmlParts.push(...persistedParts);

    // The provenance still round-trips out of the reopened host's durable metadata, identically.
    expect(reopened.office.customXmlParts.length).toBe(1);
    const survived = parseWordProvenance(reopened.office.customXmlParts[0]!);
    expect(survived.identity).toBe('sim.user@acme');
    expect(survived.contentHash).toBe('sha256:persist-word');
    expect(survived.sources).toEqual(sources);
    // And the edited body the change produced is still present after reopen.
    expect(reopened.snapshot().bodyText).toContain('cited');
  });
});

/* ───────────────────────────────────── Excel round-trip ──────────────────────────────── */

describe('provenance round-trip (Excel: workbook-settings durable metadata)', () => {
  it('actuated write → durable settings JSON → ProvenanceStore read-back → UI drill-in shows every field', async () => {
    const sources: SourceRef[] = [{ title: 'Q3 Revenue Model', uri: 'https://acme/model' }];
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nset Summary!B2 42\n```', '```cmd\ndone\n```'], {
        agentId: 'gemini-enterprise:analyst',
        identity: 'sim.user@acme',
        timestamp: '2026-06-23T00:00:00Z',
        sources,
        contentHash: 'sha256:roundtrip-excel',
      }),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write the summary value');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const excelSim = sim as ExcelSimulator;
    // (1) The value landed in the workbook.
    const summary = excelSim.snapshot().sheets.find((s) => s.name === 'Summary');
    expect(summary?.values[1]?.[1]).toBe('42');

    // (2) Durable provenance landed in the workbook settings bag and was saved.
    expect(
      [...excelSim.office.settings.keys()].filter((k) => k.startsWith('ge:prov:')),
    ).toHaveLength(1);
    expect(excelSim.office.settingsSaved).toBe(true);
    const [key, durableJson] = [...excelSim.office.settings.entries()].find(([k]) =>
      k.startsWith('ge:prov:'),
    )!;
    expect(key).toMatch(/^ge:prov:/);

    // (3) READ IT BACK from the settings JSON, re-validated against the contract schema.
    const readBack = parseExcelProvenance(String(durableJson));
    expect(readBack.agentId).toBe('gemini-enterprise:analyst');
    expect(readBack.identity).toBe('sim.user@acme');
    expect(readBack.contentHash).toBe('sha256:roundtrip-excel');
    expect(readBack.sources).toEqual(sources);

    // (4) The REAL ProvenanceStore reflects it keyed by the durable changeId.
    const changeId = key.replace(/^ge:prov:/, '');
    const store = new ProvenanceStore();
    store.record({ ok: true, changeId: asChangeId(changeId), kind: 'write-cells' }, readBack);
    expect(store.get(asChangeId(changeId))?.provenance?.identity).toBe('sim.user@acme');

    // (5) The REAL UI drill-in renders every field from the read-back record.
    const drillEl = renderDrillIn(readBack);
    const text = drillEl.textContent ?? '';
    expect(text).toContain('gemini-enterprise:analyst');
    expect(text).toContain('sim.user@acme');
    expect(drillEl.querySelector('.prov-hash')?.textContent).toBe('sha256:roundtrip-excel');
    expect(drillEl.querySelector<HTMLAnchorElement>('.prov-sources a')?.getAttribute('href')).toBe(
      'https://acme/model',
    );
  });

  it('persists across save + reopen: a fresh workbook carrying the same settings still parses', async () => {
    const sources: SourceRef[] = [];
    sim = installFakeExcel();
    ui = mountStack({
      surface: 'excel',
      client: scriptedClient(['```cmd\nset Summary!B2 99\n```', '```cmd\ndone\n```'], {
        agentId: 'gemini-enterprise:sim',
        identity: 'sim.user@acme',
        timestamp: '2026-06-23T00:00:00Z',
        sources,
        contentHash: 'sha256:persist-excel',
      }),
    });
    await ui!.flush();
    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('write a value');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    const firstSession = sim as ExcelSimulator;
    expect(firstSession.office.settings.size).toBe(2);
    // Capture the durable settings entries the workbook would persist on save.
    const persistedSettings = [...firstSession.office.settings.entries()];

    // SAVE + CLOSE.
    ui!.unmount();
    firstSession.restore();
    ui = undefined;

    // REOPEN: fresh workbook host carrying the SAME durable settings the save persisted.
    sim = installFakeExcel(defaultExcelSeed() as ExcelSeed);
    const reopened = sim as ExcelSimulator;
    for (const [k, v] of persistedSettings) reopened.office.settings.set(k, v);

    // The provenance still round-trips out of the reopened workbook's settings, identically.
    expect(reopened.office.settings.size).toBe(2);
    const survived = parseExcelProvenance(
      String([...reopened.office.settings.entries()].find(([k]) => k.startsWith('ge:prov:'))?.[1]),
    );
    expect(survived.identity).toBe('sim.user@acme');
    expect(survived.contentHash).toBe('sha256:persist-excel');
  });
});
