// @vitest-environment jsdom
/**
 * Live feature scenarios: real Gemini Enterprise + the real pane stack over simulated Office
 * documents. Off by default (it calls the engine); run with `bun run scenarios`.
 *
 *   GE_SCENARIOS=1                       enable
 *   GE_SCENARIO=excel-format,word-reply  only these ids (substring match)
 *   GE_SCENARIO_SURFACE=excel             only this app's scenarios (the release gate runs one per app)
 *   GE_SCENARIO_RUN_DIR=<dir>             write results here instead of a new timestamped folder
 *   GE_SCENARIO_SHARD=0/4                 only every 4th scenario, starting at the first
 *   GE_SCENARIO_IDS=id1,id2               only these exact ids (the gate's daily sample)
 *
 * Gemini Enterprise allows 160 assistant queries a day per licence, pooled with real users. The
 * first quota 429 writes QUOTA_EXHAUSTED next to the run folder; every later scenario is skipped
 * (not run, not failed) instead of spending what is left of the pool.
 */
import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterAll, describe, expect, it } from 'vitest';
import { runScenario, type Scenario, type ScenarioResult } from './runner.js';

const LIVE = process.env.GE_SCENARIOS === '1';
const REPO_ROOT = join(__dirname, '..', '..', '..', '..');
const DIR = join(REPO_ROOT, 'scenarios');
const filter = (process.env.GE_SCENARIO ?? '')
  .split(',')
  .map((s) => s.trim())
  .filter(Boolean);

const exactIds = process.env.GE_SCENARIO_IDS
  ? new Set(process.env.GE_SCENARIO_IDS.split(',').map((s) => s.trim()))
  : undefined;

function loadScenarios(): Scenario[] {
  return (
    readdirSync(DIR)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .flatMap((f) => JSON.parse(readFileSync(join(DIR, f), 'utf8')) as Scenario[])
      // quality-baseline.json is an object, not scenarios.
      .filter((s) => typeof s?.id === 'string')
      .filter((s) => !filter.length || filter.some((id) => s.id.includes(id)))
      .filter((s) => !exactIds || exactIds.has(s.id))
      .filter((s) => !surfaceFilter || s.surface === surfaceFilter)
      .filter((_, i) => i % shard.count === shard.index)
  );
}

/** `GE_SCENARIO_SHARD=<index>/<count>` runs every count-th scenario, so shards can run in parallel. */
const shard = (() => {
  const raw = process.env.GE_SCENARIO_SHARD?.trim();
  if (!raw) return { index: 0, count: 1 };
  const m = /^(\d+)\/(\d+)$/.exec(raw);
  const index = Number(m?.[1]);
  const count = Number(m?.[2]);
  // A bad value must fail loudly, not silently run nothing.
  if (!m || count < 1 || index >= count)
    throw new Error(`GE_SCENARIO_SHARD=${raw}: want i/n, 0 ≤ i < n`);
  return { index, count };
})();

const surfaceFilter = process.env.GE_SCENARIO_SURFACE?.trim();
const scenarios = LIVE ? loadScenarios() : [];
const results: ScenarioResult[] = [];
const runDir =
  process.env.GE_SCENARIO_RUN_DIR ??
  join(DIR, '.results', new Date().toISOString().replace(/[:.]/g, '-'));
/** Shared by every shard and app of one run (the gate passes each app its own subfolder). */
const quotaFile = join(runDir, '..', 'QUOTA_EXHAUSTED');
/** Written when Gemini refused a whole scenario (auth, outage): later ones are not run either. */
const refusedFile = join(runDir, '..', 'GEMINI_REFUSED');

describe.skipIf(!LIVE)('live scenarios', () => {
  for (const scenario of scenarios) {
    it(
      `${scenario.id}: ${scenario.title ?? scenario.prompt}`,
      async (ctx) => {
        if (existsSync(quotaFile) || existsSync(refusedFile)) ctx.skip();
        const result = await runScenario(scenario);
        if (result.notRun) {
          // Not a model or pane result: record it as NOT RUN and stop spending the shared pool.
          results.push(result);
          mkdirSync(runDir, { recursive: true });
          writeFileSync(join(runDir, `${scenario.id}.json`), JSON.stringify(result, null, 2));
          writeFileSync(
            refusedFile,
            `${new Date().toISOString()} ${scenario.id} ${result.notRun}\n`,
          );
          ctx.skip();
        }
        if (result.errors.some((e) => e.includes('(429)'))) {
          writeFileSync(quotaFile, `${new Date().toISOString()} ${scenario.id}\n`);
          ctx.skip();
        }
        results.push(result);
        mkdirSync(runDir, { recursive: true });
        writeFileSync(join(runDir, `${scenario.id}.json`), JSON.stringify(result, null, 2));
        expect(result.failures, result.failures.join('\n')).toEqual([]);
      },
      (scenario.timeoutMs ?? 240_000) + 30_000,
    );
  }

  afterAll(() => {
    if (!results.length) return;
    const rows = results.map(
      (r) =>
        `| ${r.notRun ? 'NOT RUN' : r.passed ? 'PASS' : 'FAIL'} | ${r.id} | ${r.status ?? '-'} | ${
          Object.entries(r.applied)
            .map(([k, n]) => `${k}×${n}`)
            .join(', ') || '-'
        } | ${(r.durationMs / 1000).toFixed(1)}s | ${r.notRun ?? (r.failures.join('; ') || '-')} |`,
    );
    const summary = [
      `# Scenario run ${new Date().toISOString()}`,
      '',
      `${results.filter((r) => r.passed).length}/${results.filter((r) => !r.notRun).length} passed` +
        (results.some((r) => r.notRun)
          ? ` · ${results.filter((r) => r.notRun).length} not run (Gemini refused)`
          : ''),
      '',
      '| Result | Scenario | Status | Applied | Time | Failures |',
      '| --- | --- | --- | --- | --- | --- |',
      ...rows,
      '',
    ].join('\n');
    mkdirSync(runDir, { recursive: true });
    writeFileSync(join(runDir, 'summary.md'), summary);
    console.log(`\n${summary}\nDetails: ${runDir}`);
  });
});
