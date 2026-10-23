/** Live release smoke: GE_LIVE_MODEL_SELECTION=1 bunx vitest run this file. */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import { parsePlanBlock, parseProgramBlock, type AssistRequest } from '@ge/contracts';
import { StreamAssistClient, type GeminiSkillRoute } from '@ge/gemini-client';
import { shellConfigFromEnv } from '../taskpane/config.js';
import { googleAccessToken, liveFetch, loadShellEnv, type TranscriptEntry } from './live-env.js';

const live = process.env.GE_LIVE_MODEL_SELECTION === '1';
const dir =
  process.env.GE_MODEL_SELECTION_RUN_DIR ??
  join('scenarios', '.results', `model-selection-${Date.now()}`);
const prompts: Record<GeminiSkillRoute, string> = {
  default:
    'The selected text states that revenue is 42 million. What revenue does it state? Answer in one sentence with the number.',
  planner:
    'Plan a read-only summary of the selected text. Reply with one closed ```plan block: intent summarize, surface word, scope selection, one step summarize the revenue, confidence high.',
  command:
    'Reply with one closed ```cmd block containing only echo "42" and done. Do not write to the document.',
  document:
    'The selected text states that revenue is 42 million. Summarize it in one sentence with the number.',
};

describe.skipIf(!live)('live model selection', () => {
  let client: StreamAssistClient;
  const transcript: TranscriptEntry[] = [];
  beforeAll(() => {
    const env = loadShellEnv();
    // Prompt-shape rules are pane concerns; this smoke checks only model routing.
    const { assistantRules: _rules, ...config } = shellConfigFromEnv({
      ...env,
      VITE_GE_MODEL_ID: '',
    });
    client = new StreamAssistClient(
      { getAccessToken: async () => googleAccessToken() },
      config,
      liveFetch(env, transcript),
    );
    mkdirSync(dir, { recursive: true });
    writeFileSync(join(dir, 'frozen-prompts.json'), JSON.stringify(prompts, null, 2));
  });

  for (const route of ['default', 'planner', 'command'] as const) {
    for (const modelId of [undefined, 'gemini-4-argon']) {
      it(`${modelId ?? 'engine-default'} ${route}`, async () => {
        const id = `${modelId ?? 'engine-default'}-${route}`;
        client.setModelId(modelId);
        const request: AssistRequest = {
          intent: 'ask',
          query: prompts[route],
          unit: {
            connectors: [],
            surfaceContext: { kind: 'word', selection: 'Revenue is 42 million.' },
          },
        };
        const events = [];
        const start = transcript.length;
        try {
          for await (const event of client.stream(request, { skillRoute: route }))
            events.push(event);
        } finally {
          writeFileSync(
            join(dir, `${id}.json`),
            JSON.stringify(
              { modelId: modelId ?? null, route, events, transcript: transcript.slice(start) },
              null,
              2,
            ),
          );
        }
        const calls = transcript.slice(start);
        expect(calls.length).toBeGreaterThan(0);
        for (const call of calls) {
          expect(call.status).toBe(200);
          expect(JSON.parse(call.request).generationSpec).toEqual(
            modelId ? { modelId } : undefined,
          );
        }
        expect(events.filter((event) => event.type === 'error')).toEqual([]);
        const text = events
          .flatMap((event) => (event.type === 'token' ? [event.text] : []))
          .join('');
        if (route === 'default') expect(text).toContain('42');
        if (route === 'planner') {
          const parsed = parsePlanBlock(text);
          expect(parsed.errors).toEqual([]);
          expect(parsed.plan?.intent).toBe('summarize');
        }
        if (route === 'command') {
          const parsed = parseProgramBlock(text);
          expect(parsed.found).toBe(true);
          expect(parsed.entries.length).toBeGreaterThan(0);
          expect(text).toContain('42');
        }
      }, 120_000);
    }
  }
});
