/**
 * Runs one scenario the way a user would: a seeded (simulated) Office document, the real pane stack
 * (runtime extensions → composeSession → PanelController), the real Gemini Enterprise engine, and
 * every approval gate accepted. Returns what happened plus the failed expectations.
 */
import { commandPaletteFor, intentsForManifest, type Surface } from '@ge/contracts';
import type { DocBridge, RunRecord } from '@ge/runtime';
import {
  inferImplicitIntent,
  invocationToGrounding,
  isActuating,
  shouldUsePlannerForFreeText,
} from '../taskpane/components/App.js';
import { parseComposerInput } from '../taskpane/components/Composer.js';
import { invocationToSeed } from '../taskpane/components/quick-action-seed.js';
import { extractDirectCommandProgram } from '../taskpane/direct-command.js';
import { composeSession } from '../compose.js';
import { PanelController, type PanelState } from '../controller.js';
import { connectPanelRuntime } from '../panel-runtime.js';
import { createApplicationRuntime } from '../runtime-extensions.js';
import { shellConfigFromEnv } from '../taskpane/config.js';
import { selectBridge } from '../taskpane/select-bridge.js';
import { initialUnit } from '../taskpane/unit.js';
import { excelSeed, installFakeExcel } from '../test-harness/fake-excel.js';
import { installFakeOutlook, outlookSeed } from '../test-harness/fake-outlook.js';
import { installFakePowerPoint, powerPointSeed } from '../test-harness/fake-powerpoint.js';
import { installFakeWord, wordSeed } from '../test-harness/fake-word.js';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  liveFetch,
  loadShellEnv,
  replayFetch,
  scenarioAuth,
  scenarioBrand,
  type TranscriptEntry,
} from './live-env.js';

/** The model replies of a recorded run (`GE_SCENARIO_REPLAY=<results dir>`), in order. */
function recordedReplies(dir: string, id: string): string[] {
  const run = JSON.parse(readFileSync(join(dir, `${id}.json`), 'utf8')) as {
    transcript?: Array<{ reply?: string }>;
  };
  return (run.transcript ?? []).map((entry) => entry.reply ?? '');
}

export interface ScenarioExpect {
  /** Final status of the last task (default "completed"; "any" skips the check). */
  status?: RunRecord['status'] | 'any';
  /** No error step and no panel error (default true). */
  noErrors?: boolean;
  /** Minimum number of applied changes per kind, e.g. { "format-cells": 1 }. */
  applied?: Record<string, number>;
  /** At least one change of ANY of these kinds (for vague asks with several valid answers). */
  appliedAny?: string[];
  /** Kinds that must not be applied. */
  notApplied?: string[];
  /** Case-insensitive text the final answer must contain. */
  answerIncludes?: string[];
  /** Excel: format properties at least one formatted range must carry (e.g. "fontColor"). */
  formatsInclude?: string[];
  /** Case-insensitive text that must appear somewhere in the document afterwards. */
  documentIncludes?: string[];
  /** Case-insensitive text that must NOT appear in the document afterwards. */
  documentExcludes?: string[];
  /** PowerPoint: exact one-based positions of new slides, compared with the seed's stable IDs. */
  insertedSlidePositions?: number[];
  /** PowerPoint: inspect actual table cells under the requested column headers. */
  tableColumns?: Record<string, { allowed?: string[]; includes?: string[]; nonEmpty?: boolean }>;
  tableRowCount?: number;
  /** Rows in matching tables: strings match substrings; arrays whitelist exact cell values. */
  tableRows?: Array<Record<string, string | string[]>>;
  /**
   * Excel: border edge → weight that one formatted range must carry, e.g. { "EdgeTop": "Thick" }.
   * A uniform thin grid must not pass "thick outside" (WS5 T2, Argon mx-xl-format-10).
   */
  bordersInclude?: Record<string, string>;
  /**
   * Other outcomes that are equally correct: the run passes when it meets `expect` OR any one of
   * these (e.g. a reply comment, or a clarifying question when the data has no answer).
   */
  alternatives?: Omit<ScenarioExpect, 'alternatives'>[];
}

export interface Scenario {
  id: string;
  surface: Exclude<Surface, 'onenote' | 'teams'>;
  title?: string;
  prompt: string;
  /** The document before the prompt; shape depends on the surface (see scenarios/README.md). */
  seed: Record<string, unknown>;
  /** What to answer if the planner asks a clarifying question. */
  clarification?: string;
  /** Approve plans and writes (default true). Set false to check that nothing is written unapproved. */
  approve?: boolean;
  expect?: ScenarioExpect;
  timeoutMs?: number;
}

/**
 * Why a scenario never reached the model, or undefined when any Gemini call got an answer. A run
 * refused by quota, auth or an outage says nothing about the model or the pane: it is NOT RUN,
 * never FAIL (WS5 T3: a quota outage read as 23 PowerPoint failures, 2026-10-19).
 * A local command program needs no Gemini call; its document checks still decide pass or fail.
 */
export function geminiRefusal(
  transcript: readonly TranscriptEntry[],
  localProgram = false,
): string | undefined {
  if (transcript.length === 0) return localProgram ? undefined : 'no Gemini call was made';
  // Active/failed HTTP 200 streams are executed runs; their task/deadline checks decide the verdict.
  if (transcript.some((t) => t.status === 200 && t.bodyState && t.bodyState !== 'complete'))
    return undefined;
  if (transcript.some((t) => t.status === 200 && t.reply.trim() !== '')) return undefined;
  const refused = [...new Set(transcript.filter((t) => t.status !== 200).map((t) => t.status))];
  return refused.length > 0
    ? `Gemini refused every call (HTTP ${refused.join(', ')})`
    : 'Gemini answered nothing (HTTP 200, empty)';
}

export interface ScenarioResult {
  id: string;
  surface: string;
  prompt: string;
  passed: boolean;
  failures: string[];
  durationMs: number;
  status?: RunRecord['status'];
  applied: Record<string, number>;
  failedEffects: Array<{ kind: string; errorCode?: string }>;
  answer: string;
  errors: string[];
  gates: string[];
  steps: Array<{ kind: string; text: string }>;
  runs: RunRecord[];
  transcript: TranscriptEntry[];
  document: unknown;
  /** Set when Gemini never answered (see {@link geminiRefusal}): not a pass and not a failure. */
  notRun?: string;
}

type Sim = { snapshot(): unknown; restore(): void };

function installHost(scenario: Scenario): Sim {
  // Simulators mutate their seed. Preserve the declared before-state for outcome checks/replays.
  const seed = structuredClone(scenario.seed) as never;
  switch (scenario.surface) {
    case 'excel':
      return installFakeExcel(excelSeed(seed));
    case 'word':
      return installFakeWord(wordSeed(seed));
    case 'powerpoint':
      return installFakePowerPoint(powerPointSeed(seed));
    case 'outlook':
      return installFakeOutlook(outlookSeed(seed));
  }
}

/** Plain JSON for the report (Maps become objects). */
function plain(value: unknown): unknown {
  return JSON.parse(
    JSON.stringify(value, (_k, v: unknown) => (v instanceof Map ? Object.fromEntries(v) : v)),
  );
}

const pending = (s: PanelState): boolean =>
  Boolean(
    s.pendingPlanClarification ||
    s.pendingCommandPlan ||
    s.pendingPlan ||
    s.pendingWrite ||
    s.pendingShare,
  );

export async function runScenario(scenario: Scenario): Promise<ScenarioResult> {
  const started = Date.now();
  const transcript: TranscriptEntry[] = [];
  const gates: string[] = [];
  const sim = installHost(scenario);
  const runtime = createApplicationRuntime();
  let dispose = (): void => {};
  try {
    const env = loadShellEnv();
    const bridge = selectBridge(scenario.surface);
    if (!bridge) throw new Error(`no bridge for ${scenario.surface}`);
    const { session } = await composeSession({
      // The selected brand (brands/current), as the web build applies it: its name and writing
      // rules reach every turn, so a scenario sees what users of that build see.
      config: shellConfigFromEnv(env, scenarioBrand()),
      auth: scenarioAuth,
      bridge,
      unit: initialUnit({ surface: scenario.surface }),
      hooks: runtime.hooks,
      triggers: runtime.triggers,
      primeOnHostEvent: false,
      discoverCatalog: false,
      fetchImpl: process.env.GE_SCENARIO_REPLAY
        ? replayFetch(recordedReplies(process.env.GE_SCENARIO_REPLAY, scenario.id), transcript)
        : liveFetch(env, transcript),
    });
    const controller = new PanelController(session, bridge);
    const panel = connectPanelRuntime({ session, bridge, controller, triggers: runtime.triggers });
    panel.start();
    dispose = () => {
      panel.dispose();
      controller.cancel();
      session.dispose();
    };
    const runsBefore = session.executions.list().length;

    // Act on every gate the pane would show, exactly as a user clicking through would.
    const approve = scenario.approve !== false;
    // Click like a person: on a later tick (never inside the state notification that staged the
    // gate), one gate at a time, then wait before treating the panel as idle.
    let lastClick = 0;
    let scheduled = false;
    const act = (): void => {
      scheduled = false;
      const s = controller.getState();
      if (s.pendingPlanClarification) {
        gates.push('clarification');
        controller.answerPlanClarification(scenario.clarification ?? 'Use sensible defaults.');
      } else if (s.pendingCommandPlan) {
        gates.push('confirm-plan');
        if (approve) controller.confirmCommandPlan();
        else controller.cancelCommandPlan();
      } else if (s.pendingPlan) {
        gates.push('approve-plan');
        if (approve) controller.approvePlan();
        else controller.rejectPlan();
      } else if (s.pendingWrite) {
        gates.push('approve-write');
        if (approve) controller.approvePendingWrite();
        else controller.rejectPendingWrite();
      } else if (s.pendingShare) {
        gates.push('approve-share');
        if (approve) controller.approvePendingShare();
        else controller.rejectPendingShare();
      } else return;
      lastClick = Date.now();
    };
    const unsubscribe = controller.subscribe((s) => {
      if (pending(s) && !scheduled) {
        scheduled = true;
        setTimeout(act, 300);
      }
    });

    const timeoutMs = scenario.timeoutMs ?? 240_000;
    await dispatchLikeComposer(controller, bridge, scenario);
    // Settle: idle (not busy, no gate) for 2 s in a row, or the time budget runs out.
    let idleSince = 0;
    while (Date.now() - started < timeoutMs) {
      const s = controller.getState();
      if (pending(s) && !scheduled) {
        scheduled = true; // a gate staged without a notification we caught
        setTimeout(act, 300);
      }
      if (!s.busy && !pending(s) && Date.now() - lastClick > 3000) {
        idleSince ||= Date.now();
        if (Date.now() - idleSince >= 2000) break;
      } else idleSince = 0;
      await new Promise((r) => setTimeout(r, 250));
    }
    unsubscribe();

    const state = controller.getState();
    const runs = session.executions.list().slice(runsBefore);
    const effects = runs.flatMap((r) => r.effects);
    const applied: Record<string, number> = {};
    for (const e of effects) if (e.ok) applied[e.kind] = (applied[e.kind] ?? 0) + 1;
    const answer = [...state.messages].reverse().find((m) => m.role === 'assistant')?.text ?? '';
    const errors = [
      ...(state.error ? [state.error] : []),
      ...state.steps.filter((s) => s.kind === 'error').map((s) => s.text),
    ];
    const document = plain(sim.snapshot());
    const result: ScenarioResult = {
      id: scenario.id,
      surface: scenario.surface,
      prompt: scenario.prompt,
      passed: false,
      failures: [],
      durationMs: Date.now() - started,
      ...(runs.length ? { status: runs.at(-1)!.status } : {}),
      applied,
      failedEffects: effects
        .filter((e) => !e.ok)
        .map((e) => ({ kind: e.kind, ...(e.errorCode ? { errorCode: e.errorCode } : {}) })),
      answer,
      errors: [...new Set(errors)],
      gates,
      steps: state.steps.map((s) => ({ kind: s.kind, text: s.text })),
      runs,
      // The recorder may still settle after cancellation: freeze the observed evidence and verdict.
      transcript: transcript.map((entry) => ({ ...entry })),
      document,
    };
    if (Date.now() - started >= timeoutMs) result.failures.push(`timed out after ${timeoutMs} ms`);
    result.failures.push(...checkScenario(scenario, result));
    result.passed = result.failures.length === 0;
    const refusal = geminiRefusal(
      result.transcript,
      result.runs.some((run) => run.mode === 'program'),
    );
    if (refusal) {
      result.notRun = refusal;
      result.passed = false;
    }
    return result;
  } finally {
    dispose();
    runtime.dispose();
    sim.restore();
  }
}

/**
 * Route the prompt exactly as the pane's composer does on Enter (App `dispatch`): pasted CLI runs
 * directly, actions go to the planner first, anything else is chat. Same helpers, same order.
 */
async function dispatchLikeComposer(
  controller: PanelController,
  bridge: DocBridge,
  scenario: Scenario,
): Promise<void> {
  const allowedIntents = intentsForManifest(await bridge.getCapabilities());
  const palette = commandPaletteFor(scenario.surface, allowedIntents);
  const inv = parseComposerInput(scenario.prompt, { kind: 'selection' }, palette);
  const program = extractDirectCommandProgram(inv.raw);
  if (program) return controller.runDirectCommands(program);
  const intent = inferImplicitIntent(scenario.surface, allowedIntents, inv);
  const routed =
    intent !== undefined && intent !== inv.intent
      ? { ...inv, intent, raw: `/${intent} ${inv.raw}`.trim() }
      : inv;
  const seed = invocationToSeed(routed);
  const grounding = invocationToGrounding(inv);
  const planFirst =
    isActuating(intent) ||
    (intent === undefined && shouldUsePlannerForFreeText(allowedIntents, inv));
  if (planFirst) return controller.proposePlan(seed, grounding);
  if (isActuating(intent)) return controller.runCommands(seed, grounding);
  return controller.send(seed, grounding);
}

export function checkScenario(scenario: Scenario, r: ScenarioResult): string[] {
  const main = checkExpect(scenario.expect ?? {}, r, scenario.seed);
  if (main.length === 0) return main;
  const alternatives = scenario.expect?.alternatives ?? [];
  return alternatives.some((alt) => checkExpect(alt, r, scenario.seed).length === 0) ? [] : main;
}

function checkExpect(e: ScenarioExpect, r: ScenarioResult, seed: Scenario['seed']): string[] {
  const out: string[] = [];
  const status = e.status ?? 'completed';
  if (status !== 'any' && r.status !== status)
    out.push(`status is ${r.status ?? 'none'}, expected ${status}`);
  if ((e.noErrors ?? true) && r.errors.length) out.push(`errors shown: ${r.errors.join(' | ')}`);
  for (const [kind, min] of Object.entries(e.applied ?? {})) {
    if ((r.applied[kind] ?? 0) < min)
      out.push(`applied ${r.applied[kind] ?? 0} × ${kind}, expected ≥ ${min}`);
  }
  if (e.appliedAny?.length && !e.appliedAny.some((kind) => (r.applied[kind] ?? 0) > 0)) {
    out.push(`none of ${e.appliedAny.join(', ')} was applied`);
  }
  for (const kind of e.notApplied ?? []) {
    if (r.applied[kind]) out.push(`${kind} was applied ${r.applied[kind]} time(s), expected none`);
  }
  const answer = r.answer.toLowerCase();
  for (const text of e.answerIncludes ?? []) {
    if (!answer.includes(text.toLowerCase())) out.push(`answer does not mention "${text}"`);
  }
  if (e.formatsInclude?.length) {
    const formats = Object.values(
      (r.document as { formats?: Record<string, Record<string, unknown>> }).formats ?? {},
    );
    for (const facet of e.formatsInclude) {
      if (!formats.some((f) => f[facet] !== undefined)) out.push(`no range was given ${facet}`);
    }
  }
  if (e.bordersInclude) {
    const formats = Object.values(
      (
        r.document as {
          formats?: Record<string, { borders?: Record<string, { weight?: string }> }>;
        }
      ).formats ?? {},
    );
    const wanted = Object.entries(e.bordersInclude);
    if (
      !formats.some((f) => wanted.every(([edge, weight]) => f.borders?.[edge]?.weight === weight))
    )
      out.push(
        `no range has borders ${wanted.map(([edge, weight]) => `${edge}=${weight}`).join(', ')}`,
      );
  }
  const doc = JSON.stringify(r.document).toLowerCase();
  for (const text of e.documentIncludes ?? []) {
    if (!doc.includes(text.toLowerCase())) out.push(`document does not contain "${text}"`);
  }
  for (const text of e.documentExcludes ?? []) {
    if (doc.includes(text.toLowerCase()))
      out.push(`document contains "${text}", expected it absent`);
  }
  if (e.tableColumns) {
    const normal = (text: string) => text.trim().toLowerCase();
    const slides =
      (r.document as { slides?: Array<{ shapes?: Array<{ table?: string[][] }> }> }).slides ?? [];
    const tables = slides
      .flatMap((slide) => slide.shapes ?? [])
      .flatMap((shape) => (shape.table ? [shape.table] : []));
    const columns = Object.entries(e.tableColumns);
    const matching = tables.filter((table) =>
      columns.every(([name]) => table[0]?.some((cell) => normal(cell) === normal(name))),
    );
    if (!matching.length)
      out.push(`no table has columns ${columns.map(([name]) => name).join(', ')}`);
    for (const table of matching) {
      if (e.tableRowCount !== undefined && table.length - 1 !== e.tableRowCount)
        out.push(`table has ${table.length - 1} data rows, expected ${e.tableRowCount}`);
      const remaining = table.slice(1);
      for (const expected of e.tableRows ?? []) {
        const index = remaining.findIndex((row) =>
          Object.entries(expected).every(([name, wanted]) => {
            const column = table[0]!.findIndex((cell) => normal(cell) === normal(name));
            if (column < 0) return false;
            const cell = normal(row[column] ?? '');
            return typeof wanted === 'string'
              ? cell.includes(normal(wanted))
              : wanted.map(normal).includes(cell);
          }),
        );
        if (index < 0) out.push(`table is missing expected row ${JSON.stringify(expected)}`);
        else remaining.splice(index, 1);
      }
      for (const [name, rules] of columns) {
        const index = table[0]!.findIndex((cell) => normal(cell) === normal(name));
        const cells = table.slice(1).map((row) => normal(row[index] ?? ''));
        if (rules.nonEmpty && (!cells.length || cells.some((cell) => !cell)))
          out.push(`table column ${name} has empty cells`);
        if (rules.allowed && cells.some((cell) => !rules.allowed!.map(normal).includes(cell)))
          out.push(`table column ${name} contains an unsupported value`);
        for (const text of rules.includes ?? []) {
          if (!cells.some((cell) => cell.includes(normal(text))))
            out.push(`table column ${name} does not contain "${text}"`);
        }
      }
    }
  }
  if (e.insertedSlidePositions) {
    const before = seed.slides as Array<{ id?: string }> | undefined;
    const after = (r.document as { slides?: Array<{ id?: string }> }).slides;
    if (!before || !after || [...before, ...after].some((slide) => !slide.id)) {
      out.push('slide positions require seed and result slides with stable IDs');
    } else {
      const original = new Set(before.map((slide) => slide.id));
      const retained = after.filter((slide) => original.has(slide.id)).map((slide) => slide.id);
      if (JSON.stringify(retained) !== JSON.stringify(before.map((slide) => slide.id)))
        out.push('original slides were removed or reordered');
      const positions = after.flatMap((slide, i) => (original.has(slide.id) ? [] : [i + 1]));
      if (JSON.stringify(positions) !== JSON.stringify(e.insertedSlidePositions))
        out.push(
          `new slides are at ${JSON.stringify(positions)}, expected ${JSON.stringify(e.insertedSlidePositions)}`,
        );
    }
  }
  return out;
}
