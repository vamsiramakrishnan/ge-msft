import { mkdtemp, readFile, rm, truncate, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { afterEach, describe, expect, it } from 'vitest';

const root = resolve(fileURLToPath(new URL('.', import.meta.url)), '../../..');
const head = 'a'.repeat(40);
let directory: string | undefined;
afterEach(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
  directory = undefined;
});
function run(args: string[]) {
  return spawnSync('bun', ['scripts/office-lab.ts', ...args], {
    cwd: root,
    encoding: 'utf8',
    timeout: 10000,
    maxBuffer: 1024 * 1024,
  });
}
describe('acceptance CLI evidence handling', () => {
  it('creates an empty exact-head campaign and returns BLOCKED_EXTERNAL rather than passing it', async () => {
    directory = await mkdtemp(join(tmpdir(), 'ge-campaign-'));
    const init = run(['campaign-init', head, 'word,excel']);
    expect(init.status).toBe(0);
    const bundle = JSON.parse(init.stdout);
    expect(bundle.evidence.native).toEqual([]);
    const file = join(directory, 'bundle.json');
    await writeFile(file, init.stdout);
    const result = run(['campaign', file, head]);
    expect(result.status).toBe(2);
    expect(JSON.parse(result.stdout)).toMatchObject({
      status: 'BLOCKED_EXTERNAL',
      failures: [],
      evidenceOnly: true,
    });
  });
  it('refuses oversized files before parsing and never repeats document contents from malformed JSON', async () => {
    directory = await mkdtemp(join(tmpdir(), 'ge-campaign-'));
    const file = join(directory, 'report.json');
    await writeFile(file, '');
    await truncate(file, 9 * 1024 * 1024);
    const large = run(['campaign', file, head]);
    expect(large.status).toBe(1);
    expect(large.stderr).toContain('eight MiB');
    await writeFile(file, 'PRIVATE_DOCUMENT_MARKER');
    const invalid = run(['campaign', file, head]);
    expect(invalid.status).toBe(1);
    expect(invalid.stderr).toContain('not valid JSON');
    expect(invalid.stderr).not.toContain('PRIVATE_DOCUMENT_MARKER');
  });
  it('collects actual typed local reports but rejects arbitrary bodies and fabricated field names', async () => {
    directory = await mkdtemp(join(tmpdir(), 'ge-campaign-'));
    const init = run(['campaign-init', head, 'word']);
    const plan = join(directory, 'plan.json');
    await writeFile(plan, JSON.stringify(JSON.parse(init.stdout).plan));
    const report = join(directory, 'local.json');
    await writeFile(
      report,
      JSON.stringify({
        version: 1,
        commit: head,
        local: [{ id: 'tests', commit: head, status: 'passed' }],
      }),
    );
    const collected = run(['campaign-collect', plan, report]);
    expect(collected.status).toBe(0);
    expect(JSON.parse(collected.stdout).evidence.local).toHaveLength(1);
    const before = await readFile(report, 'utf8');
    await writeFile(report, JSON.stringify({ providerScope: 'PRIVATE_DOCUMENT_MARKER' }));
    const rejected = run(['campaign-collect', plan, report]);
    expect(rejected.status).toBe(1);
    expect(rejected.stderr).not.toContain('PRIVATE_DOCUMENT_MARKER');
    expect(before).toContain('tests');
  });
  it('lists canonical fixture expectations without executing a host or model', () => {
    const result = run(['fixtures', 'excel']);
    expect(result.status).toBe(0);
    const fixtures = JSON.parse(result.stdout);
    expect(fixtures).toHaveLength(5);
    expect(
      fixtures.find((fixture: { id: string }) => fixture.id === 'excel-pivot-readback')
        .expectedEffects,
    ).toEqual(['write-cells', 'insert-pivot']);
  });
});
