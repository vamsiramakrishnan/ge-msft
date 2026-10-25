import { describe, expect, it } from 'vitest';
import { assessOfficeReadiness, type OfficeReleaseEvidence } from './office-readiness.js';

const commit = 'a'.repeat(40);
function evidence(): OfficeReleaseEvidence {
  return {
    version: 1,
    commit,
    requiredSurfaces: ['word'],
    profiles: [
      { id: 'word-windows', surface: 'word', platform: 'PC', build: '16.0.18000', mode: 'write' },
    ],
    local: [
      'typecheck',
      'tests',
      'lint',
      'build',
      'generated',
      'secret-scan',
      'security-review',
    ].map((id) => ({
      id: id as OfficeReleaseEvidence['local'][number]['id'],
      commit,
      status: 'passed',
    })),
    native: [
      {
        version: 1,
        fixtureVersion: 1,
        fixtureId: 'word-tracked-edit',
        commit,
        surface: 'word',
        host: {
          platform: 'PC',
          build: '16.0.18000',
          mode: 'write',
          requirements: { 'WordApi 1.4': true },
        },
        startedAt: '2026-10-25T00:00:00Z',
        durationMs: 10,
        status: 'passed',
        checks: { execution: true, postcondition: true, finished: true },
        nativeSyncs: 2,
        effects: [{ kind: 'tracked-change', ok: true, verification: 'verified' }],
      },
    ],
    tenant: [
      {
        commit,
        profileId: 'word-windows',
        status: 'passed',
        scenario: 'contract-review',
        transport: 'widget',
        skillsVerified: true,
        sourceFreshnessVerified: true,
        approvalRefusalVerified: true,
      },
    ],
    outcomes: {
      commit,
      profileIds: ['word-windows'],
      workflowIds: ['contract-review'],
      attempts: 20,
      firstPassSuccesses: 19,
      wrongTargets: 0,
      unauthorizedEffects: 0,
      falseCompletions: 0,
      baselineSyncs: 100,
      candidateSyncs: 50,
      baselineInferenceBytes: 2000,
      candidateInferenceBytes: 1000,
      artifactsMatch: true,
      freshnessRefusalsMatch: true,
    },
  };
}
describe('Office release readiness', () => {
  it('passes only matching current-head evidence for the advertised native host/build', () => {
    expect(assessOfficeReadiness(evidence(), commit)).toMatchObject({
      status: 'PASSED',
      certifiedProfiles: ['word-windows'],
      evidenceOnly: true,
    });
  });
  it('keeps unavailable external execution distinct from implementation failure', () => {
    const input = evidence();
    input.native = [];
    input.tenant = [];
    delete input.outcomes;
    const report = assessOfficeReadiness(input, commit);
    expect(report.status).toBe('BLOCKED_EXTERNAL');
    expect(report.failures).toEqual([]);
    expect(report.missing).toHaveLength(3);
    expect(report.certifiedProfiles).toEqual([]);
  });
  it('rejects wrong commit, host build and mode as certification of the advertised profile', () => {
    for (const field of ['commit', 'build', 'mode'] as const) {
      const input = evidence();
      if (field === 'commit') input.native[0]!.commit = 'b'.repeat(40);
      else input.native[0]!.host![field] = field === 'mode' ? 'read' : '16.0.17000';
      expect(assessOfficeReadiness(input, commit).status).toBe('BLOCKED_EXTERNAL');
    }
    expect(assessOfficeReadiness(evidence(), 'b'.repeat(40)).status).toBe('FAIL');
  });
  it('does not let a pass mask a failed report, incomplete postcondition or fake host', () => {
    const input = evidence();
    input.native.push({ ...input.native[0]!, status: 'failed' });
    expect(assessOfficeReadiness(input, commit).status).toBe('FAIL');
    input.native = [input.native[0]!];
    input.native[0]!.checks.postcondition = false;
    expect(assessOfficeReadiness(input, commit).status).toBe('FAIL');
    input.profiles[0]!.build = 'test-only';
    expect(assessOfficeReadiness(input, commit).status).toBe('FAIL');
  });
  it('fails safety regressions, insufficient improvement and false tenant success', () => {
    const input = evidence();
    input.outcomes!.wrongTargets = 1;
    input.outcomes!.candidateSyncs = 51;
    input.outcomes!.firstPassSuccesses = 18;
    input.tenant[0]!.approvalRefusalVerified = false;
    const report = assessOfficeReadiness(input, commit);
    expect(report.status).toBe('FAIL');
    expect(report.failures).toHaveLength(4);
  });
  it('requires the flagship scenario and matching expected native effect', () => {
    const input = evidence();
    input.tenant[0]!.scenario = 'unrelated-smoke';
    expect(assessOfficeReadiness(input, commit).status).toBe('BLOCKED_EXTERNAL');
    input.tenant[0]!.scenario = 'contract-review';
    input.native[0]!.effects[0]!.kind = 'insert-text';
    expect(assessOfficeReadiness(input, commit).status).toBe('FAIL');
  });
  it('does not extend a benchmark to untested profiles or workflows', () => {
    const input = evidence();
    input.outcomes!.profileIds = ['other-profile'];
    input.outcomes!.workflowIds = ['workbook-close'];
    expect(assessOfficeReadiness(input, commit).status).toBe('BLOCKED_EXTERNAL');
  });
  it('requires every declared surface and refuses duplicate profile identities', () => {
    const input = evidence();
    delete input.requiredSurfaces;
    expect(
      assessOfficeReadiness(input, commit).missing.filter((item) => item.startsWith('Advertised')),
    ).toHaveLength(3);
    input.profiles.push({ ...input.profiles[0]! });
    expect(assessOfficeReadiness(input, commit).status).toBe('FAIL');
  });
});
