import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  commandIntentText,
  READ_VERBS,
  registryEntriesForSurface,
  type CapabilityManifest,
  type Surface,
} from '@ge/contracts';
import { discoverCommands, writeSignatures } from '../capability-catalog.js';
import { COMMAND_BOOTSTRAP_MAX_BYTES, renderCommandBootstrap } from '../command-protocol.js';

/**
 * Replays of live model traffic recorded on 2026-09-30 (docs/COMMAND-RELIABILITY.md). Each test
 * pins a deterministic part of the command path that the recorded run showed failing, using the
 * exact prompt text and model output that was captured, so a regression fails here without a model.
 */
interface LiveFixture {
  confirmedPlanTask: string;
  docState: string;
  executorPrograms: Record<string, string>;
}

const excelChart = JSON.parse(
  readFileSync(new URL('./live-2026-09-30-excel-chart.json', import.meta.url), 'utf8'),
) as LiveFixture;

function manifest(surface: Surface): CapabilityManifest {
  return {
    surface,
    contextKinds: ['range', 'sheet', 'shape', 'slide', 'comment', 'attachment'],
    reads: [...READ_VERBS],
    actuations: registryEntriesForSurface(surface).map((entry) => ({
      kind: entry.kind,
      surface,
      title: entry.title,
      reversible: true,
    })),
  };
}

const bytes = (text: string) => new TextEncoder().encode(text).byteLength;
const SURFACES: Surface[] = ['excel', 'word', 'powerpoint', 'outlook', 'onenote', 'teams'];

describe('A — the executor sees the exact syntax of every write it may emit', () => {
  it.each(SURFACES)('%s: every advertised write is disclosed with its exact syntax', (surface) => {
    const current = manifest(surface);
    const prompt = renderCommandBootstrap(current, excelChart.confirmedPlanTask);
    const signatures = writeSignatures(current);
    expect(signatures.length).toBeGreaterThan(0);
    expect(new Set(signatures).size).toBe(signatures.length);
    for (const line of signatures) expect(prompt).toContain(line);
    // A specialized command must show real parameters, never the uninformative placeholder.
    for (const line of signatures.filter((l) => l.startsWith('/')))
      expect(line).not.toContain('[key=value ...]');
    expect(bytes(prompt)).toBeLessThanOrEqual(COMMAND_BOOTSTRAP_MAX_BYTES);
  });

  it('shows the Excel chart signature for the recorded planner hand-off', () => {
    // Live: `chart <` never appeared in the executor prompt; the model invented four syntaxes.
    const prompt = renderCommandBootstrap(manifest('excel'), excelChart.confirmedPlanTask);
    expect(prompt).toContain(
      'chart <column|bar|line|pie|scatter|area> <range> [title="…"] [series=rows|columns]',
    );
  });
});

describe('A — relevant-command cards are ranked on the user request, not the plan wrapper', () => {
  it('extracts only the request, steps and exclusions from the recorded confirmed-plan task', () => {
    const intent = commandIntentText(excelChart.confirmedPlanTask);
    expect(intent).toContain('Create a bar chart of Total by Product');
    expect(intent).not.toMatch(/read live host content|open Microsoft 365 surface/);
  });

  it('ranks chart first for the recorded hand-off (the raw task picks read and open)', () => {
    const current = manifest('excel');
    const raw = discoverCommands(current, excelChart.confirmedPlanTask).map((c) => c.command);
    expect(raw.slice(0, 2)).toEqual(['read', 'open']); // the recorded failure, reproduced
    const ranked = discoverCommands(current, commandIntentText(excelChart.confirmedPlanTask));
    expect(ranked[0]?.command).toBe('chart');
    expect(renderCommandBootstrap(current, excelChart.confirmedPlanTask)).toContain(
      'Relevant command:\nCommand: chart',
    );
  });
});
