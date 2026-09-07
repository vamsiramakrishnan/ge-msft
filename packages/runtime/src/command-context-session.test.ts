import { describe, expect, it, vi } from 'vitest';
import { asChangeId, type CapabilityManifest } from '@ge/contracts';
import {
  CommandContextSession,
  type CommandContextRenderInput,
  type CommandContextRuntimeState,
} from './command-context-session.js';
import { CommandCapsule, CommandCapsuleBudgetError } from './command-capsule.js';
import { ExecutionState } from './execution-state.js';
import { CommandResultStore } from './result-store.js';
import { renderCommandBootstrap, renderGrammarPrompt } from './command-protocol.js';

const capabilities: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['range'],
  reads: ['read'],
  actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
};
const task = 'Reconcile totals; preserve currency and write once.';
const state: CommandContextRuntimeState = {
  analysisBindings: [['source', 'a_0123456789abcdef01234567']],
  composeBindings: new Map([['total', { kind: 'number' as const, value: 42 }]]),
  artifacts: [],
  effects: [
    {
      changeId: asChangeId('landed-once'),
      kind: 'write-cells',
      ok: true,
      recoveryPending: true,
      verification: { status: 'unknown' },
    },
  ],
  externalShareAttempts: 0,
};
const input: CommandContextRenderInput = {
  capabilities,
  docState: '<doc_state>CURRENT</doc_state>',
  skills: [{ name: 'inspect_total', params: ['range'], body: ['read $range'] }],
  state: () => state,
};
const normalizeScopes = (query: string): string =>
  query.replace(/\b(state|result):[0-9a-z]+:/g, '$1:SCOPE:');
const metrics = () => ({
  resultInputBytes: 0,
  resultInputBytesComplete: true,
  resultOutputBytes: 0,
});

// These compare the previous public renderers with the facade, including exact framing/JSON bytes.
describe.each(['projection', 'transcript'] as const)(
  'CommandContextSession %s compatibility',
  (mode) => {
    it.each(['compact', 'full'] as const)(
      'preserves %s wire requests, corrections and encoded metrics',
      (disclosure) => {
        const context = new CommandContextSession({ contextMode: mode, disclosure });
        context.begin(task);
        const previous =
          mode === 'projection' ? new ExecutionState(task) : new CommandCapsule(task);
        const previousResults = new CommandResultStore({ inlineBytes: 4096 });
        const counters = metrics();
        const protocol =
          disclosure === 'full'
            ? renderGrammarPrompt(capabilities)
            : renderCommandBootstrap(capabilities, task);
        const expectedInput = {
          protocol,
          docState: input.docState,
          skills: input.skills,
          ...(mode === 'projection'
            ? {
                bindings: [
                  {
                    name: 'source',
                    kind: 'artifact' as const,
                    value: { id: state.analysisBindings[0]![1], available: false },
                  },
                  { name: 'total', kind: 'number' as const, value: { kind: 'number', value: 42 } },
                ],
                artifacts: [],
                effects: state.effects,
                externalShareAttempts: 0,
              }
            : {}),
        };
        expect(normalizeScopes(context.render(input))).toBe(
          normalizeScopes(previous.render(expectedInput)),
        );
        const results = [{ text: '42' }, { error: 'Inspect previous write before retrying.' }];
        const encoded = previousResults.encode(results);
        context.record({ program: 'read Data!A1', results }, counters);
        const observation = { program: 'read Data!A1', resultsJson: encoded.text };
        if (previous instanceof ExecutionState) previous.append(observation, results);
        else previous.append(observation);
        expect(normalizeScopes(context.render(input))).toBe(
          normalizeScopes(previous.render(expectedInput)),
        );
        expect(counters).toEqual({
          resultInputBytes: encoded.inputBytes,
          resultInputBytesComplete: encoded.inputBytesComplete,
          resultOutputBytes: encoded.outputBytes,
        });
        context.record(
          { program: 'Unfenced response', correction: 'Use one closed cmd fence.' },
          counters,
        );
        previous.append({ program: 'Unfenced response', correction: 'Use one closed cmd fence.' });
        expect(normalizeScopes(context.render(input))).toBe(
          normalizeScopes(previous.render(expectedInput)),
        );
        expect(counters.resultInputBytes).toBe(encoded.inputBytes);
        expect(context.snapshotPolicy).toBe('fresh');
      },
    );
  },
);

describe('command context lifecycle', () => {
  it('preserves conversation framing and avoids a snapshot or state collection on corrections', () => {
    const context = new CommandContextSession({ sessionMode: 'conversation' });
    const collectState = vi.fn(input.state!);
    context.begin(task);
    expect(context.isolated).toBe(false);
    expect(context.snapshotPolicy).toBe('deduplicate');
    expect(context.render({ ...input, state: collectState })).toBe(
      [renderCommandBootstrap(capabilities, task), input.docState, `TASK:\n${task}`, 'Begin.'].join(
        '\n\n',
      ),
    );
    context.record({ program: 'read Data!A1', results: [{ text: '42' }] });
    expect(context.render(input)).toBe(
      '```result\n[{"text":"42"}]\n```\n\n<doc_state>CURRENT</doc_state>\n\n(Continue. Next command?)',
    );
    context.record({ program: 'No fence', correction: 'Use one closed cmd fence.' });
    expect(context.snapshotPolicy).toBe('none');
    expect(context.render({ ...input, state: collectState })).toBe('Use one closed cmd fence.');
    expect(collectState).not.toHaveBeenCalled();
    context.record({ program: 'read Data!A1', results: [] });
    expect(context.snapshotPolicy).toBe('deduplicate');
    expect(context.render({ capabilities })).toBe(
      '```result\n[]\n```\n\n(Continue. Next command?)',
    );
  });

  it('collects live state only for projection and snapshots caller mode options', () => {
    const options = { contextMode: 'transcript' as 'projection' | 'transcript' };
    const context = new CommandContextSession(options);
    options.contextMode = 'projection';
    context.begin(task);
    const collectState = vi.fn(input.state!);
    context.render({ ...input, state: collectState });
    expect(collectState).not.toHaveBeenCalled();
  });

  it('encodes one result batch once and expires both namespaces together on begin/clear', () => {
    const context = new CommandContextSession();
    const counters = metrics();
    context.begin(task);
    context.record({ program: 'read Data!A1', results: [{ text: 'x'.repeat(5000) }] }, counters);
    const rendered = context.render({ capabilities });
    const resultRef = (
      JSON.parse(/```result\n([\s\S]*?)\n```/.exec(rendered)![1]!) as Array<{ ref: string }>
    )[0]!.ref;
    const execution = JSON.parse(
      /<execution_state[^>]*>\n([\s\S]*?)\n<\/execution_state>/.exec(rendered)![1]!,
    ) as { latest: { receipt: string } };
    const stateRef = execution.latest.receipt;
    expect(context.inspect(`${resultRef} path=/text limit=1`)).toMatchObject({ preview: 'x' });
    expect(context.inspect(`${stateRef} path=/program`)).toMatchObject({ preview: 'read Data!A1' });
    const recordedCounters = { ...counters };
    expect(context.render({ capabilities })).toBe(rendered);
    expect(counters).toEqual(recordedCounters);
    context.begin('New task');
    expect(context.inspect(resultRef)).toMatchObject({ code: 'reference' });
    expect(context.inspect(stateRef)).toMatchObject({ code: 'reference' });
    context.clear();
    expect(context.inspect(stateRef)).toMatchObject({ code: 'reference' });
    expect(context.inspect('Data!A1')).toBeUndefined();
    expect(() => context.render({ capabilities })).toThrow('Begin a command context');
    expect(() => context.record({ program: 'read', results: [] })).toThrow(
      'Begin a command context',
    );
  });

  it('keeps encoding counters truthful when a later journal append fails', () => {
    const context = new CommandContextSession({ contextMode: 'transcript', maxBytes: 128 });
    context.begin('Read');
    const counters = metrics();
    const results = [{ text: 'observed' }];
    const encoded = new CommandResultStore({ inlineBytes: 4096 }).encode(results);
    expect(() => context.record({ program: 'x'.repeat(200), results }, counters)).toThrow(
      CommandCapsuleBudgetError,
    );
    expect(counters).toEqual({
      resultInputBytes: encoded.inputBytes,
      resultInputBytesComplete: true,
      resultOutputBytes: encoded.outputBytes,
    });
  });

  it('leaves no active context after an oversized begin fails', () => {
    const context = new CommandContextSession({ maxBytes: 128 });
    expect(() => context.begin('x'.repeat(129))).toThrow(CommandCapsuleBudgetError);
    expect(() => context.render({ capabilities })).toThrow('Begin a command context');
  });
});
