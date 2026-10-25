import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  CAPABILITY_REGISTRY,
  asChangeId,
  approvalClassOf,
  isReversibleKind,
  parseCommandLine,
  type ActuationResult,
  type HostCapabilityProfile,
  type Surface,
} from '@ge/contracts';
import type { CommandLoopEvent, PlanEffect, RunCommandsOptions } from './assist-session.js';
import { compileCommand, isCompileError } from './command-protocol.js';
import { nativeTraces } from './native-trace.js';
import {
  certifyNativeFixture,
  nativeFixture,
  nativeFixtureExpectation,
  nativeOperationFixture,
  NATIVE_OPERATION_FIXTURE_IDS,
  type NativeFixture,
} from './native-certification.js';

const commit = 'a'.repeat(40);
/** Synthetic profiles and streams test validation only; they establish no native certification. */
function profile(surface: Surface): HostCapabilityProfile {
  return {
    surface,
    platform: 'PC',
    build: '16.0.validator-fixture',
    mode: surface === 'outlook' ? 'compose' : 'write',
    requirements: Object.fromEntries(
      CAPABILITY_REGISTRY.filter((entry) => entry.surface === surface)
        .flatMap((entry) => entry.requirementSets)
        .map((set) => [`${set.name}:${set.minVersion}`, true]),
    ),
  };
}
function bridge(fixture: NativeFixture) {
  let key = 'saved-native-document';
  let host = profile(fixture.surface);
  return {
    surface: fixture.surface,
    documentKey: () => key,
    getHostProfile: () => host,
    switchDocument: () => {
      key = 'different-document';
    },
    setProfile: (next: HostCapabilityProfile) => {
      host = next;
    },
    readRange: vi.fn(async () => [
      {
        ref: {
          id: 'slide',
          surface: fixture.surface,
          kind: 'slide' as const,
          title: 'Native fixture',
        },
        value: { as: 'text' as const, text: fixture.expectedText ?? 'no marker' },
      },
    ]),
  };
}
function effectsFor(program: string, start: number): PlanEffect[] {
  return program
    .split('\n')
    .filter((line) => line && !/^(finish|done)\b/.test(line))
    .map((line, index) => {
      const parsed = parseCommandLine(line);
      if ('error' in parsed) throw new Error(parsed.error);
      const surface: Surface =
        /^(?:grid|table)$/.test(line.split(' ')[0]!) ||
        /^\/(?:insert-pivot|filter-range|set-data-validation)/.test(line)
          ? 'excel'
          : /^\/(?:apply-slide-layout)/.test(line) || /^slide /.test(line)
            ? 'powerpoint'
            : /^(?:compose|\/(?:set-subject|set-body))\b/.test(line)
              ? 'outlook'
              : 'word';
      const compiled = compileCommand(parsed, {
        surface,
        mintChangeId: () => asChangeId(`fixture-${start + index}`),
      });
      if (isCompileError(compiled) || compiled.kind !== 'write')
        throw new Error('Fixture did not compile to a host effect: ' + line);
      return {
        request: compiled.request,
        command: line,
        approvalClass: approvalClassOf(compiled.request.kind),
        reversible: isReversibleKind(compiled.request.kind),
      };
    });
}
function simulatedSession(
  fixture: NativeFixture,
  transform: (receipt: ActuationResult, index: number) => ActuationResult = (receipt) => receipt,
) {
  let dispatched = 0;
  const calls: string[] = [];
  const session = {
    async *runCommandProgram(
      program: string,
      options: RunCommandsOptions,
    ): AsyncGenerator<CommandLoopEvent> {
      calls.push(program);
      const effects = effectsFor(program, dispatched);
      yield {
        type: 'plan-preview',
        turn: 1,
        effects,
        dag: [],
        approvalClasses: [...new Set(effects.map((effect) => effect.approvalClass))],
      };
      const approved = await options.approvePlan?.(effects);
      for (const effect of effects) {
        if (approved) await options.beforeDispatch?.();
        const request = effect.request;
        const receipt: ActuationResult = approved
          ? {
              ok: true,
              changeId: request.changeId,
              kind: request.kind,
              location: 'slide:1',
              ...(fixture.verifiedEffects?.includes(dispatched)
                ? { verification: { status: 'verified' as const } }
                : {}),
              ...(fixture.reviewOpened ? { completion: 'review-opened' as const } : {}),
            }
          : {
              ok: false,
              changeId: request.changeId,
              kind: request.kind,
              error: { code: 'plan_unapproved', message: 'No approval.' },
            };
        yield {
          type: 'write-result',
          turn: 1,
          changeId: request.changeId,
          result: transform(receipt, dispatched++),
        };
      }
      yield { type: 'done', turn: 1, answer: '' };
    },
  };
  return { session, calls };
}
const layoutCapture = {
  version: 1 as const,
  documentId: 'sha256:' + 'a'.repeat(64),
  complete: true,
  omissions: [],
  masters: [
    {
      id: 'master-1',
      name: 'Theme',
      layouts: [
        { id: 'layout-1', name: 'Title', revision: 'sha256:' + 'b'.repeat(64), placeholders: [] },
      ],
    },
  ],
};
function operationFixtures(): NativeFixture[] {
  return [
    nativeOperationFixture({ id: 'word-table-readback', disposable: true }),
    nativeOperationFixture({
      id: 'excel-table-readback',
      disposable: true,
      sourceRange: "'FY 26'!A1:B3",
    }),
    nativeOperationFixture({
      id: 'excel-pivot-readback',
      disposable: true,
      sourceRange: 'Source!A1:B3',
      destinationRange: 'Output!A1:C6',
    }),
    nativeOperationFixture({
      id: 'excel-filter-readback',
      disposable: true,
      sourceRange: 'Source!A1:B3',
    }),
    nativeOperationFixture({
      id: 'excel-validation-readback',
      disposable: true,
      targetRange: 'Target!A1:A3',
    }),
    nativeOperationFixture({
      id: 'powerpoint-layout-readback',
      disposable: true,
      capture: layoutCapture,
      slideId: 'slide-1',
      masterId: 'master-1',
      layoutId: 'layout-1',
    }),
    nativeOperationFixture({ id: 'outlook-compose-readback', disposable: true }),
  ];
}
afterEach(() => {
  vi.useRealTimers();
});

describe('native certification boundary (simulated streams)', () => {
  it('reports unavailable without calling the executor for missing, unknown or simulated hosts', async () => {
    for (const host of [
      undefined,
      { ...profile('excel'), platform: 'simulator' },
      { ...profile('excel'), build: 'unknown' },
    ]) {
      const runCommandProgram = vi.fn();
      const report = await certifyNativeFixture(
        { runCommandProgram },
        { surface: 'excel', documentKey: () => 'saved', getHostProfile: () => host },
        nativeFixture('excel', 'Sheet1!A1:B2'),
        commit,
        {},
      );
      expect(report.status).toBe('unavailable');
      expect(report.effects).toEqual([]);
      expect(runCommandProgram).not.toHaveBeenCalled();
    }
  });
  it('requires native requirement support, saved identity and the right host mode before any mutation', async () => {
    const fixture = nativeOperationFixture({ id: 'outlook-compose-readback', disposable: true });
    const runCommandProgram = vi.fn();
    for (const [host, key] of [
      [{ ...profile('outlook'), mode: 'read' as const }, 'saved'],
      [{ ...profile('outlook'), requirements: {} }, 'saved'],
      [profile('outlook'), undefined],
    ] as const) {
      const report = await certifyNativeFixture(
        { runCommandProgram },
        { surface: 'outlook', getHostProfile: () => host, documentKey: () => key },
        fixture,
        commit,
        {},
      );
      expect(report.status).toBe('unavailable');
    }
    expect(runCommandProgram).not.toHaveBeenCalled();
  });
  it('requires exact commit, an unchanged canonical fixture and an exact disposable Excel rectangle', async () => {
    expect(() => nativeFixture('excel')).toThrow('explicit');
    expect(() => nativeFixture('excel', 'A1\nset B1 999')).toThrow('explicit');
    expect(() => nativeFixture('excel', 'Sheet1!A1:C3')).toThrow('2×2');
    const fixture = nativeFixture('word');
    expect(() => {
      fixture.program += '\n/insert-text text="extra"';
    }).toThrow();
    await expect(
      certifyNativeFixture(
        { runCommandProgram: vi.fn() },
        { surface: 'word' },
        { ...fixture, program: fixture.program + '\n/insert-text text="extra"' },
        commit,
        {},
      ),
    ).rejects.toThrow('unchanged');
    const fresh = nativeFixture('word');
    await expect(
      certifyNativeFixture(
        { runCommandProgram: vi.fn() },
        { surface: 'word' },
        fresh,
        'dirty-head',
        {},
      ),
    ).rejects.toThrow('exact');
  });
  it('runs seeded native cases as separately approved programs and compares every expected effect', async () => {
    for (const fixture of operationFixtures()) {
      const { session, calls } = simulatedSession(fixture);
      const approvePlan = vi.fn(async () => true);
      const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
        approvePlan,
      });
      expect(report.status, fixture.id).toBe('passed');
      expect(report.effects.map((effect) => effect.kind)).toEqual(fixture.expectedEffects);
      expect(calls).toHaveLength(fixture.setupProgram ? 2 : 1);
      expect(approvePlan).toHaveBeenCalledTimes(calls.length);
      expect(report.checks).toMatchObject({
        approved: true,
        correlated: true,
        expectedEffects: true,
        scopeStable: true,
        hostStable: true,
        postcondition: true,
      });
      expect(JSON.stringify(report)).not.toContain('GE_CERT_');
      expect(JSON.stringify(report)).not.toContain('saved-native-document');
    }
  });
  it('keeps canonical fixture intent frozen across capability and approval awaits', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const program = fixture.program;
    const { session, calls } = simulatedSession(fixture);
    const host = {
      ...bridge(fixture),
      getCapabilities: async () => {
        await Promise.resolve();
        expect(Reflect.set(fixture, 'program', '/set-body text="bad"')).toBe(false);
        expect(Reflect.set(fixture.expectedEffects, 0, 'set-body')).toBe(false);
        return {
          surface: fixture.surface,
          contextKinds: [],
          actuations: fixture.expectedEffects.map((kind) => ({
            kind,
            surface: fixture.surface,
            title: kind,
            reversible: false,
          })),
        };
      },
    };
    const report = await certifyNativeFixture(session, host, fixture, commit, {
      approvePlan: async () => {
        await Promise.resolve();
        expect(Reflect.set(fixture, 'expectedKind', 'set-body')).toBe(false);
        expect(Reflect.set(fixture.modes, 0, 'read')).toBe(false);
        expect(Reflect.set(fixture.verifiedEffects!, 0, 99)).toBe(false);
        return true;
      },
    });
    expect(report.status).toBe('passed');
    expect(calls).toEqual([program]);
  });
  it('detaches the observed host profile before asynchronous capability discovery', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const current = profile('excel');
    const host = {
      ...bridge(fixture),
      getHostProfile: () => current,
      getCapabilities: async () => {
        await Promise.resolve();
        current.mode = 'read';
        return {
          surface: fixture.surface,
          contextKinds: [],
          actuations: fixture.expectedEffects.map((kind) => ({
            kind,
            surface: fixture.surface,
            title: kind,
            reversible: false,
          })),
        };
      },
    };
    const { session } = simulatedSession(fixture);
    const report = await certifyNativeFixture(session, host, fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.effects).toEqual([]);
    expect(report.host?.mode).toBe('write');
  });
  it('reports failed native capability discovery as unavailable without execution fallback', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const runCommandProgram = vi.fn();
    const report = await certifyNativeFixture(
      { runCommandProgram },
      {
        ...bridge(fixture),
        getCapabilities: async () => {
          throw new Error('Native discovery failed.');
        },
      },
      fixture,
      commit,
      {},
    );
    expect(report.status).toBe('unavailable');
    expect(report.checks.support).toBe(false);
    expect(runCommandProgram).not.toHaveBeenCalled();
  });
  it('keeps fallback per-write approvals sequential rather than replacing an active approval card', async () => {
    const fixture = nativeFixture('word');
    const { session } = simulatedSession(fixture);
    let active = false;
    const approveWrite = vi.fn(async () => {
      expect(active).toBe(false);
      active = true;
      await Promise.resolve();
      active = false;
      return true;
    });
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approveWrite,
    });
    expect(report.status).toBe('passed');
    expect(approveWrite).toHaveBeenCalledTimes(2);
  });
  it('does not proceed to the native object mutation if seed approval is missing or declined', async () => {
    const fixture = operationFixtures()[2]!;
    const { session, calls } = simulatedSession(fixture);
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approvePlan: () => false,
    });
    expect(report.status).toBe('failed');
    expect(calls).toHaveLength(1);
    expect(report.effects.map((effect) => effect.kind)).toEqual(['write-cells']);
    expect(report.checks.approved).toBe(false);
  });
  it('rejects a successful but mismatched receipt instead of borrowing its verification', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const { session } = simulatedSession(fixture, (receipt) => ({
      ...receipt,
      changeId: asChangeId('foreign-change'),
    }));
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.checks.correlated).toBe(false);
  });
  it('rejects an empty successful stream without inferring a host effect', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const session = {
      async *runCommandProgram(): AsyncGenerator<CommandLoopEvent> {
        yield { type: 'done', turn: 1, answer: '' };
      },
    };
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.effects).toEqual([]);
    expect(report.checks.expectedEffects).toBe(false);
  });
  it('rejects extra unrelated effects even when the expected receipt was verified', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const original = simulatedSession(fixture).session;
    const session = {
      async *runCommandProgram(
        program: string,
        options: RunCommandsOptions,
      ): AsyncGenerator<CommandLoopEvent> {
        yield* original.runCommandProgram(program, options);
        yield {
          type: 'write-result',
          turn: 1,
          changeId: 'foreign-extra',
          result: {
            ok: true,
            changeId: asChangeId('foreign-extra'),
            kind: 'set-body',
            verification: { status: 'verified' },
          },
        };
      },
    };
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.checks.correlated).toBe(false);
    expect(report.checks.expectedEffects).toBe(false);
  });
  it('fails when a landed edit cannot restore tracking mode, even if readback itself verified', async () => {
    const fixture = nativeFixture('word');
    const { session } = simulatedSession(fixture, (receipt, index) =>
      index === 1 ? { ...receipt, trackingModeRestoreFailed: true } : receipt,
    );
    const report = await certifyNativeFixture(session, bridge(fixture), fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.checks.execution).toBe(false);
    expect(report.checks.postcondition).toBe(true);
  });
  it('blocks a document switch during approval before dispatch', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const host = bridge(fixture);
    const { session } = simulatedSession(fixture);
    const report = await certifyNativeFixture(session, host, fixture, commit, {
      approvePlan: () => {
        host.switchDocument();
        return true;
      },
    });
    expect(report.status).toBe('failed');
    expect(report.effects).toEqual([]);
  });
  it('does not infer a native slide postcondition from a generic callback or an oversized read', async () => {
    const fixture = nativeFixture('powerpoint');
    const host = bridge(fixture);
    host.readRange.mockResolvedValue([
      {
        ref: { id: 'slide', surface: 'powerpoint', kind: 'slide', title: 'Native' },
        value: { as: 'text', text: 'x'.repeat(65537) + fixture.expectedText },
      },
    ]);
    const { session } = simulatedSession(fixture);
    const report = await certifyNativeFixture(session, host, fixture, commit, {
      approvePlan: () => true,
    });
    expect(report.status).toBe('failed');
    expect(report.checks.postcondition).toBe(false);
  });
  it('bounds a stalled final readback without replaying a successful fixture write', async () => {
    vi.useFakeTimers();
    const fixture = nativeFixture('powerpoint');
    const host = bridge(fixture);
    host.readRange.mockImplementation(() => new Promise(() => undefined));
    const { session, calls } = simulatedSession(fixture);
    const pending = certifyNativeFixture(session, host, fixture, commit, {
      approvePlan: () => true,
    });
    await vi.advanceTimersByTimeAsync(5001);
    const report = await pending;
    expect(report.status).toBe('failed');
    expect(report.effects[0]?.ok).toBe(true);
    expect(calls).toHaveLength(1);
  });
  it('counts observed native syncs beyond the diagnostics ring capacity', async () => {
    const fixture = nativeFixture('excel', 'Sheet1!A1:B2');
    const { session } = simulatedSession(fixture);
    const report = await certifyNativeFixture(
      session,
      bridge(fixture),
      fixture,
      commit,
      { approvePlan: () => true },
      (event) => {
        if (event.type !== 'write-result') return;
        for (let index = 0; index < 300; index++)
          nativeTraces.record({
            id: `native-fixture-${index}`,
            surface: 'excel',
            stage: 'batch',
            syncs: 1,
            syncMs: 0,
            durationMs: 0,
            outcome: 'completed',
          });
      },
    );
    expect(report.status).toBe('passed');
    expect(report.nativeSyncs).toBe(300);
    expect(nativeTraces.list().length).toBeLessThanOrEqual(200);
  });
  it('shares detached authoritative case expectations with imported-report validation', () => {
    const fixtures = operationFixtures();
    expect(fixtures.map((fixture) => fixture.id)).toEqual([...NATIVE_OPERATION_FIXTURE_IDS]);
    for (const fixture of fixtures)
      expect(nativeFixtureExpectation(fixture.id)).toEqual(
        expect.objectContaining({
          surface: fixture.surface,
          expectedEffects: fixture.expectedEffects,
        }),
      );
    expect(nativeFixtureExpectation('__proto__')).toBeUndefined();
    const copy = nativeFixtureExpectation('excel-pivot-readback')!;
    (copy.expectedEffects as string[]).push('set-body');
    expect(nativeFixtureExpectation('excel-pivot-readback')?.expectedEffects).toEqual([
      'write-cells',
      'insert-pivot',
    ]);
  });
  it('refuses unsafe fixture destinations and incomplete native layouts before command construction', () => {
    expect(() =>
      nativeOperationFixture({
        id: 'excel-table-readback',
        disposable: true,
        sourceRange: 'A1:B3',
      }),
    ).toThrow('explicit');
    expect(() =>
      nativeOperationFixture({
        id: 'excel-pivot-readback',
        disposable: true,
        sourceRange: 'Source!A1:B3',
        destinationRange: 'Source!D1:F6',
      }),
    ).toThrow('different worksheet');
    expect(() =>
      nativeOperationFixture({
        id: 'excel-validation-readback',
        disposable: true,
        targetRange: 'Target!A1:A300',
      }),
    ).toThrow('3×1');
    expect(() =>
      nativeOperationFixture({
        id: 'powerpoint-layout-readback',
        disposable: true,
        capture: { ...layoutCapture, complete: false },
        slideId: 'slide-1',
        masterId: 'master-1',
        layoutId: 'layout-1',
      }),
    ).toThrow('complete');
  });
});
