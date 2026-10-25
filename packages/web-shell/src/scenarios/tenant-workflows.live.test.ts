// @vitest-environment jsdom
/** GE_TENANT_WORKFLOWS=1; explicit token/config only. Office hosts are always simulated here. */
import { execFileSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { configuredTenantWorkflowProvider } from './tenant-workflow-env.js';
import { runFlagshipTenantSuite } from './tenant-workflows.js';

describe.skipIf(process.env.GE_TENANT_WORKFLOWS !== '1')(
  'flagship tenant acceptance (simulated Office)',
  () => {
    it('writes exact-head metadata-only evidence and distinguishes missing tenant access', async (context) => {
      const commit = execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
      const dirty =
        execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim().length > 0;
      if (process.env.GE_TENANT_EXPECT_COMMIT && process.env.GE_TENANT_EXPECT_COMMIT !== commit)
        throw new Error('The expected acceptance commit differs from this checkout.');
      const report = await runFlagshipTenantSuite({
        commit,
        dirty,
        repetitions: Number(process.env.GE_TENANT_WORKFLOW_REPETITIONS ?? 1),
        maxModelCalls: Number(process.env.GE_TENANT_WORKFLOW_MAX_CALLS ?? 12),
        provider: configuredTenantWorkflowProvider(),
        providerScope: 'live',
      });
      const destination = resolve(
        process.env.GE_TENANT_WORKFLOW_REPORT ?? 'scenarios/.results/tenant-workflows/report.json',
      );
      mkdirSync(dirname(destination), { recursive: true, mode: 0o700 });
      writeFileSync(destination, JSON.stringify(report, null, 2) + '\n', { mode: 0o600 });
      console.log(
        JSON.stringify({
          status: report.status,
          providerScope: report.providerScope,
          hostScope: report.hostScope,
          ...report.summary,
        }),
      );
      expect(report.status).not.toBe('FAIL');
      if (report.status === 'BLOCKED_EXTERNAL') context.skip();
    }, 1900000);
  },
);
