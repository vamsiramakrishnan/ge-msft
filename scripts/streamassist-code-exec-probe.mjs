#!/usr/bin/env node
// @ts-check
/**
 * Probe Gemini Enterprise streamAssist code-execution behavior.
 *
 * This is a live API probe, not a unit test. It records whether the assistant
 * actually streams AssistantContent.executableCode / codeExecutionResult parts,
 * and separately scores whether the final answer is correct.
 *
 * Standard endpoint env:
 *   GE_TOKEN or GE_ACCESS_TOKEN
 *   GE_ENDPOINT, or GE_PROJECT + GE_LOCATION + GE_ENGINE
 *
 * Widget endpoint env:
 *   GE_TOKEN or GE_ACCESS_TOKEN
 *   GE_WIDGET_CONFIG_ID
 *   GE_WIDGET_SERVER_TOKEN
 *   GE_LOCATION=global
 *   GE_ENGINE=ge-msft-plugin-test_1782382759735   # optional, seeds widget session name
 *
 * Optional:
 *   GE_DATA_STORES=collections/default_collection/dataStores/...
 *   GE_SKILL_NAMES=projects/.../agents/...
 *   GE_USER_PROJECT=quota project
 *   GE_TIME_ZONE=Australia/Melbourne
 *
 * Usage:
 *   bun scripts/streamassist-code-exec-probe.mjs
 *   bun scripts/streamassist-code-exec-probe.mjs --widget --require-code table followup
 *   bun scripts/streamassist-code-exec-probe.mjs --dry-run
 */

import { mkdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';

const SCENARIOS = [
  {
    id: 'arithmetic',
    title: 'Pure calculation',
    expectCode: true,
    query:
      'Use code execution if available. Compute the sum of integers from 1 through 100. Return only the total and one short sentence.',
    mustMatch: [/5050/],
  },
  {
    id: 'table',
    title: 'Inline CSV aggregate',
    expectCode: true,
    query: `Use code execution if available to analyze this CSV. Compute total revenue, total cost, and gross margin = (revenue - cost) / revenue.

Item,Region,Revenue,Cost
Line 1,EMEA,120,70
Line 2,APAC,90,55
Line 3,AMER,110,60
Line 4,EMEA,1500,300
Line 5,APAC,95,50
Line 6,AMER,105,65`,
    mustMatch: [/2020/, /\b600\b/, /70(\.3|\.30|%)/],
  },
  {
    id: 'followup',
    title: 'Session carry-over from previous CSV turn',
    expectCode: false,
    query:
      'Using the CSV from the previous turn, which line has the highest revenue? Include the line name and revenue.',
    mustMatch: [/Line 4/i, /1500/],
  },
  {
    id: 'failure-shape',
    title: 'Division-by-zero handling',
    expectCode: true,
    query:
      'Use code execution if available to evaluate 1 / 0 in Python, then explain the failure in one sentence without a stack trace.',
    mustMatch: [/zero|division|fail|error/i],
  },
  {
    id: 'schedule-csv',
    title: 'Large rectangular schedule generation',
    expectCode: true,
    query:
      'Use code execution if available to generate a CSV-shaped weekly schedule with columns Time, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday. Include 30-minute rows from 08:00 through 18:00 for a Sunnyvale Google SWE who has Tuesday 08:00 music, morning India team overlap, late afternoon Australia manager overlap, daily fitness, and focused coding blocks. Return the CSV header, the first two data rows, the last data row, and the total row count.',
    mustMatch: [/Time,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday/i, /Music/i, /India/i, /Australia|manager/i, /21|twenty-one/i],
  },
];

const OUT_JSON = 'dist/probes/streamassist-code-exec.json';
const OUT_MD = 'dist/probes/streamassist-code-exec.md';

function usage() {
  console.log(`streamassist-code-exec-probe

Options:
  --widget        Use content-discoveryengine widgetStreamAssist body.
  --standard      Use official assistant :streamAssist body (default).
  --require-code  Treat missing code-execution parts as failure for scenarios that expect code.
  --dry-run       Print the scenarios and selected endpoint, then exit.

Scenario ids: ${SCENARIOS.map((s) => s.id).join(', ')}
`);
}

function parseArgs(argv) {
  const flags = new Set(argv.filter((a) => a.startsWith('--')));
  if (flags.has('--help') || flags.has('-h')) {
    usage();
    process.exit(0);
  }
  const ids = argv.filter((a) => !a.startsWith('--'));
  const selected = ids.length ? SCENARIOS.filter((s) => ids.includes(s.id)) : SCENARIOS;
  if (ids.length && selected.length !== ids.length) {
    const known = new Set(SCENARIOS.map((s) => s.id));
    const bad = ids.filter((id) => !known.has(id));
    throw new Error(`Unknown scenario id(s): ${bad.join(', ')}`);
  }
  return {
    widget: flags.has('--widget'),
    requireCode: flags.has('--require-code'),
    dryRun: flags.has('--dry-run'),
    selected,
  };
}

function token() {
  return process.env.GE_TOKEN || process.env.GE_ACCESS_TOKEN || '';
}

function standardUrl() {
  if (process.env.GE_ENDPOINT) return process.env.GE_ENDPOINT;
  const { GE_PROJECT, GE_LOCATION = 'global', GE_ENGINE } = process.env;
  if (!GE_PROJECT || !GE_ENGINE) {
    throw new Error('Set GE_ENDPOINT, or GE_PROJECT + GE_LOCATION + GE_ENGINE.');
  }
  const host =
    GE_LOCATION === 'global'
      ? 'discoveryengine.googleapis.com'
      : `discoveryengine.${GE_LOCATION}.rep.googleapis.com`;
  return (
    `https://${host}/v1alpha/projects/${GE_PROJECT}/locations/${GE_LOCATION}/collections/` +
    `default_collection/engines/${GE_ENGINE}/assistants/default_assistant:streamAssist`
  );
}

function standardUrlOrPlaceholder() {
  try {
    return standardUrl();
  } catch {
    return 'https://discoveryengine.googleapis.com/v1alpha/projects/${GE_PROJECT}/locations/${GE_LOCATION}/collections/default_collection/engines/${GE_ENGINE}/assistants/default_assistant:streamAssist';
  }
}

function widgetUrl() {
  const location = process.env.GE_LOCATION || 'global';
  return `https://content-discoveryengine.googleapis.com/v1alpha/locations/${location}/widgetStreamAssist`;
}

function dataStores() {
  return splitEnv('GE_DATA_STORES').map((dataStore) => ({ dataStore }));
}

function skills() {
  return splitEnv('GE_SKILL_NAMES').map((name) => ({ name }));
}

function splitEnv(name) {
  return (process.env[name] || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

function seededWidgetSession() {
  if (process.env.GE_WIDGET_SESSION) return process.env.GE_WIDGET_SESSION;
  if (!process.env.GE_ENGINE) return undefined;
  const suffix = String(Date.now());
  return `collections/default_collection/engines/${process.env.GE_ENGINE}/sessions/${suffix}`;
}

function buildStandardBody(query, session) {
  const body = {
    query: { text: query },
    userMetadata: { timeZone: process.env.GE_TIME_ZONE || 'UTC' },
  };
  if (session) body.session = session;
  const ds = dataStores();
  if (ds.length) body.toolsSpec = { vertexAiSearchSpec: { dataStoreSpecs: ds } };
  const sk = skills();
  if (sk.length) body.skillsSpec = { skills: sk };
  return body;
}

function buildWidgetBody(query, session) {
  const configId = process.env.GE_WIDGET_CONFIG_ID;
  if (!configId) throw new Error('--widget requires GE_WIDGET_CONFIG_ID.');
  const streamAssistRequest = {
    query: { parts: [{ text: query }] },
    answerGenerationMode: 'NORMAL',
    languageCode: 'en-US',
    userMetadata: { timeZone: process.env.GE_TIME_ZONE || 'UTC' },
    assistSkippingMode: 'REQUEST_ASSIST',
  };
  if (session) streamAssistRequest.session = session;
  const ds = dataStores();
  const toolsSpec = {};
  if (ds.length) toolsSpec.vertexAiSearchSpec = { dataStoreSpecs: ds };
  toolsSpec.webGroundingSpec = {};
  toolsSpec.toolRegistry = 'default_tool_registry';
  if (Object.keys(toolsSpec).length) streamAssistRequest.toolsSpec = toolsSpec;
  const sk = skills();
  if (sk.length) streamAssistRequest.skillsSpec = { skills: sk };
  return {
    configId,
    additionalParams: { token: '-', origin: 'ORIGIN_UNSPECIFIED' },
    streamAssistRequest,
  };
}

async function callStreamAssist({ widget, query, session }) {
  const bearer = token();
  if (!bearer) throw new Error('Set GE_TOKEN or GE_ACCESS_TOKEN.');
  const headers = {
    Authorization: `Bearer ${bearer}`,
    'Content-Type': 'application/json',
  };
  if (process.env.GE_USER_PROJECT) headers['x-goog-user-project'] = process.env.GE_USER_PROJECT;
  if (widget) {
    if (!process.env.GE_WIDGET_SERVER_TOKEN) {
      throw new Error('--widget requires GE_WIDGET_SERVER_TOKEN.');
    }
    headers['x-server-token'] = process.env.GE_WIDGET_SERVER_TOKEN;
    headers.Origin = 'https://vertexaisearch.cloud.google';
  }
  const body = widget ? buildWidgetBody(query, session) : buildStandardBody(query, session);
  const res = await fetch(widget ? widgetUrl() : standardUrl(), {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });
  const raw = await res.text();
  if (!res.ok) throw new Error(`streamAssist ${res.status}: ${raw.slice(0, 600)}`);
  return parseStreamAssistPayload(raw, session);
}

function parseStreamAssistPayload(raw, fallbackSession) {
  const chunks = parseChunks(raw);
  const events = [];
  let text = '';
  let session = fallbackSession;
  for (const chunk of chunks) {
    session = chunk?.sessionInfo?.session || chunk?.streamAssistResponse?.sessionInfo?.session || session;
    const answer = chunk?.answer || chunk?.streamAssistResponse?.answer;
    for (const reply of answer?.replies || []) {
      const content = reply?.groundedContent?.content || {};
      if (content.thought === true) {
        if (typeof content.text === 'string' && content.text.trim()) {
          events.push({ type: 'activity', text: compactActivity(content.text) });
        }
        continue;
      }
      if (content.executableCode?.code) {
        events.push({ type: 'code-execution', code: content.executableCode.code });
      }
      if (content.codeExecutionResult) {
        events.push({
          type: 'code-execution-result',
          outcome: content.codeExecutionResult.outcome,
          output: content.codeExecutionResult.output || '',
        });
      }
      if (content.text) {
        text += content.text;
        events.push({ type: 'token', text: content.text });
      }
    }
  }
  return { rawHash: sha256(raw), chunks: chunks.length, text, events, session };
}

function compactActivity(text) {
  const compact = text
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .join(' ');
  return compact.length > 160 ? `${compact.slice(0, 157)}...` : compact;
}

function parseChunks(raw) {
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch {
    return raw
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return { rawText: line };
        }
      });
  }
}

function sha256(text) {
  return createHash('sha256').update(text).digest('hex');
}

function scoreScenario(scenario, parsed, requireCode) {
  const searchable = [
    parsed.text,
    ...parsed.events.map((event) => ('output' in event ? event.output : '')),
  ].join('\n');
  const missing = scenario.mustMatch
    .filter((pattern) => !pattern.test(searchable))
    .map((pattern) => String(pattern));
  const codeEvents = parsed.events.filter(
    (event) => event.type === 'code-execution' || event.type === 'code-execution-result',
  );
  const codeObserved = codeEvents.length > 0;
  const codeMissing = scenario.expectCode && requireCode && !codeObserved;
  return {
    status: missing.length || codeMissing ? 'fail' : codeObserved || !scenario.expectCode ? 'pass' : 'pass_no_code',
    missing,
    codeObserved,
    codeEventCount: codeEvents.length,
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const endpoint = args.widget ? widgetUrl() : standardUrlOrPlaceholder();
  if (args.dryRun) {
    console.log(
      JSON.stringify(
        {
          endpoint,
          mode: args.widget ? 'widget' : 'standard',
          requireCode: args.requireCode,
          scenarios: args.selected.map(({ id, title, expectCode }) => ({ id, title, expectCode })),
        },
        null,
        2,
      ),
    );
    return;
  }

  let session = args.widget ? seededWidgetSession() : undefined;
  const results = [];
  for (const scenario of args.selected) {
    const startedAt = new Date().toISOString();
    try {
      const parsed = await callStreamAssist({
        widget: args.widget,
        query: scenario.query,
        session,
      });
      session = parsed.session || session;
      const score = scoreScenario(scenario, parsed, args.requireCode);
      results.push({
        id: scenario.id,
        title: scenario.title,
        startedAt,
        status: score.status,
        missing: score.missing,
        codeObserved: score.codeObserved,
        codeEventCount: score.codeEventCount,
        chunkCount: parsed.chunks,
        rawResponseSha256: parsed.rawHash,
        textPreview: parsed.text.slice(0, 500),
        events: parsed.events.map((event) =>
          event.type === 'code-execution'
            ? { ...event, code: `${event.code.slice(0, 500)}${event.code.length > 500 ? '\n...' : ''}` }
            : event,
        ),
      });
      console.log(
        `[${score.status}] ${scenario.id}: code=${score.codeObserved ? 'yes' : 'no'} chunks=${parsed.chunks}`,
      );
    } catch (error) {
      results.push({
        id: scenario.id,
        title: scenario.title,
        startedAt,
        status: 'error',
        message: error instanceof Error ? error.message : String(error),
      });
      console.log(`[error] ${scenario.id}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  const report = {
    generatedAt: new Date().toISOString(),
    mode: args.widget ? 'widget' : 'standard',
    endpoint,
    requireCode: args.requireCode,
    session,
    results,
    summary: {
      total: results.length,
      pass: results.filter((r) => r.status === 'pass').length,
      passNoCode: results.filter((r) => r.status === 'pass_no_code').length,
      fail: results.filter((r) => r.status === 'fail').length,
      error: results.filter((r) => r.status === 'error').length,
      codeObserved: results.filter((r) => r.codeObserved).length,
    },
  };
  await mkdir('dist/probes', { recursive: true });
  await writeFile(OUT_JSON, `${JSON.stringify(report, null, 2)}\n`);
  await writeFile(OUT_MD, renderMarkdown(report));
  console.log(`wrote ${OUT_JSON}`);
  console.log(`wrote ${OUT_MD}`);
  if (report.summary.fail || report.summary.error) process.exitCode = 1;
}

function renderMarkdown(report) {
  const rows = report.results
    .map(
      (r) =>
        `| ${r.id} | ${r.status} | ${r.codeObserved ? 'yes' : 'no'} | ${r.codeEventCount ?? 0} | ${
          r.missing?.join(', ') || r.message?.replace(/\|/g, '\\|') || ''
        } |`,
    )
    .join('\n');
  return `# streamAssist code-execution probe

- generatedAt: ${report.generatedAt}
- mode: ${report.mode}
- endpoint: ${report.endpoint}
- requireCode: ${report.requireCode}
- session: ${report.session || '(none)'}

| scenario | status | code observed | code events | notes |
|---|---:|---:|---:|---|
${rows}
`;
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
