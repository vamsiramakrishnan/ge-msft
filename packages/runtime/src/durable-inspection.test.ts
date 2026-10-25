import { describe, expect, it, vi } from 'vitest';
import {
  asChangeId,
  contentHash,
  makeCellSnapshot,
  sourceVersion,
  type ActuationRequest,
  type ActuationResult,
  type CellSnapshot,
  type ContextRef,
  type DurableCommandProgram,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import {
  deriveDurableInspection,
  inspectDurableOutput,
  validateDurableSources,
} from './durable-inspection.js';

const id = asChangeId('inspection-effect');
const doc = 'sha256:' + 'a'.repeat(64);
const program = (
  body = 'set Output!A1 "literal"\nfinish when=verified',
): DurableCommandProgram => ({
  version: 1,
  id: 'inspection-task',
  surface: 'excel',
  evidence: { kind: 'literal' },
  stages: [{ id: 'write', dependsOn: [], program: body }],
});
async function cell(
  value = 'output',
  objectId = 'sheet-output',
  locator = 'Output!A1',
): Promise<CellSnapshot> {
  return makeCellSnapshot({
    surface: 'excel',
    documentId: doc,
    objectId,
    locator,
    values: [[value]],
    formulas: [['']],
  });
}
const bridge = (
  captureCells?: DocBridge['captureCells'],
): Pick<DocBridge, 'surface' | 'documentKey' | 'captureCells' | 'captureTextTarget'> => ({
  surface: 'excel',
  documentKey: () => 'saved-workbook',
  captureCells,
});
const receipt = (kind: ActuationResult['kind'], afterHash?: string): ActuationResult => ({
  ok: true,
  kind,
  changeId: id,
  verification: { status: 'verified', ...(afterHash ? { afterHash } : {}) },
});
async function cellEvidence() {
  const before = await cell('before');
  const input: ActuationRequest = {
    surface: 'excel',
    kind: 'write-cells',
    changeId: id,
    params: { target: { range: before.locator }, cellValues: [['output']] },
    preconditions: [sourceVersion(before)],
  };
  return { before, input, evidence: deriveDurableInspection(input)! };
}

describe('read-only durable task currentness inspection', () => {
  it('captures native identity without persisting values, text or inverse bodies', async () => {
    const { evidence } = await cellEvidence();
    expect(evidence).toMatchObject({
      kind: 'cells',
      documentId: doc,
      objectId: 'sheet-output',
      locator: 'Output!A1',
    });
    expect(JSON.stringify(evidence)).not.toMatch(/before|output"\]|cellValues|text|inverse/);
    expect(
      deriveDurableInspection({
        surface: 'excel',
        kind: 'write-cells',
        changeId: id,
        params: { target: { range: 'Output!A1' }, cellValues: [['output']] },
      }),
    ).toBeUndefined();
    expect(
      deriveDurableInspection({
        surface: 'word',
        kind: 'tracked-change',
        changeId: id,
        params: { text: 'replacement' },
      }),
    ).toBeUndefined();
  });
  it('requires a correlated verified native hash and refuses legacy or unsupported checkpoints without reading', async () => {
    const { evidence } = await cellEvidence();
    const capture = vi.fn(async () => cell());
    expect(
      (await inspectDurableOutput(undefined, receipt('write-cells'), bridge(capture))).status,
    ).toBe('unsupported');
    expect(
      (await inspectDurableOutput(evidence, receipt('write-cells'), bridge(capture))).status,
    ).toBe('unsupported');
    expect(
      (
        await inspectDurableOutput(
          evidence,
          { ...receipt('write-cells', (await cell()).hash), changeId: asChangeId('foreign') },
          bridge(capture),
        )
      ).status,
    ).toBe('unsupported');
    expect(capture).not.toHaveBeenCalled();
  });
  it('checks current cell data plus native worksheet/document identity', async () => {
    const { evidence } = await cellEvidence();
    const output = await cell();
    expect(
      (
        await inspectDurableOutput(
          evidence,
          receipt('write-cells', output.hash),
          bridge(async () => output),
        )
      ).status,
    ).toBe('current');
    for (const current of [
      await cell('manual edit'),
      await cell('output', 'replacement-sheet'),
      await makeCellSnapshot({ ...output, documentId: 'foreign' }),
    ]) {
      expect(
        (
          await inspectDurableOutput(
            evidence,
            receipt('write-cells', output.hash),
            bridge(async () => current),
          )
        ).status,
      ).toBe('changed');
    }
    // A port that returns a stale hash with mutated values cannot certify this saved output.
    expect(
      (
        await inspectDurableOutput(
          evidence,
          receipt('write-cells', output.hash),
          bridge(async () => ({ ...output, values: [['edited']] })),
        )
      ).status,
    ).toBe('unsupported');
  });
  it('refuses foreign document switches and unavailable/oversized native reads', async () => {
    const { evidence } = await cellEvidence();
    const output = await cell();
    let key = 'saved';
    const switched = {
      ...bridge(async () => {
        key = 'other';
        return output;
      }),
      documentKey: () => key,
    };
    expect(
      (await inspectDurableOutput(evidence, receipt('write-cells', output.hash), switched)).status,
    ).toBe('changed');
    expect(
      (await inspectDurableOutput(evidence, receipt('write-cells', output.hash), bridge())).status,
    ).toBe('unsupported');
    const large = await makeCellSnapshot({
      surface: 'excel',
      documentId: doc,
      objectId: 'sheet-output',
      locator: 'Output!A1:AO1',
      values: [Array.from({ length: 41 }, () => 'x'.repeat(32768))],
      formulas: [Array(41).fill('')],
    });
    const oversized = { ...evidence, locator: large.locator };
    expect(
      (
        await inspectDurableOutput(
          oversized,
          receipt('write-cells', large.hash),
          bridge(async () => large),
        )
      ).status,
    ).toBe('unsupported');
  });
  it.each(['word', 'powerpoint'] as const)(
    'uses exact %s native text targets and the actual postcondition hash',
    async (surface) => {
      const text = 'Current generated text';
      const afterHash = await contentHash({ text });
      const kind = surface === 'word' ? 'fill-content-control' : 'set-shape-text';
      const target =
        surface === 'word'
          ? { contentControlId: 'cc-1' }
          : { slideId: 'slide-1', shapeId: 'shape-1' };
      const evidence = deriveDurableInspection(
        { surface, kind, changeId: id, params: { target, text } },
        doc,
      )!;
      expect(evidence).toBeDefined();
      const capture = {
        text,
        target: {
          surface,
          documentId: doc,
          objectId: surface === 'word' ? 'cc-1' : 'shape-1',
          ...(surface === 'powerpoint' ? { parentObjectId: 'slide-1' } : {}),
          locator: 'object',
          revision: afterHash,
          scope: 'native-text',
        },
      };
      const native = {
        surface,
        documentKey: () => 'saved',
        captureTextTarget: vi.fn(async (_ref: ContextRef) => capture),
      };
      expect((await inspectDurableOutput(evidence, receipt(kind, afterHash), native)).status).toBe(
        'current',
      );
      expect(native.captureTextTarget.mock.calls[0]?.[0]).toMatchObject({
        hostRef:
          surface === 'word'
            ? { type: 'word.contentControl', contentControlId: 'cc-1' }
            : { type: 'powerpoint.shape', slideId: 'slide-1', shapeId: 'shape-1' },
      });
      capture.text = 'Manual edit';
      expect((await inspectDurableOutput(evidence, receipt(kind, afterHash), native)).status).toBe(
        'changed',
      );
      expect(
        deriveDurableInspection(
          {
            surface,
            kind,
            changeId: id,
            params: { target: { ...target, expectedDocumentId: 'sha256:' + 'b'.repeat(64) }, text },
          },
          doc,
        ),
      ).toBeUndefined();
    },
  );
  it('does not infer source coverage or accept read/analysis/expression/formula programs as literal inputs', async () => {
    const noEvidence = { ...program(), evidence: undefined };
    expect((await validateDurableSources(noEvidence, bridge())).status).toBe('unsupported');
    expect((await validateDurableSources(program(), bridge())).status).toBe('current');
    for (const body of [
      'read Source!A1\nset Output!A1 "literal"\nfinish when=verified',
      'let x = read Source!A1\nset Output!A1 = $x\nfinish when=verified',
      'set Output!A1 (read Source!A1 | count)\nfinish when=verified',
      'set Output!A1 =SUM(Source!A1:A10)\nfinish when=verified',
      'set A1 "literal"\nfinish when=verified',
      'set Output!A1 "one"\nset Output!A2 "two"\nfinish when=verified',
      'done "Nothing to inspect"',
    ])
      expect((await validateDurableSources(program(body), bridge())).status).toBe('unsupported');
  });
  it('recaptures every explicitly declared source and detects drift, source replacement and forged hashes', async () => {
    const source = await cell('source', 'sheet-source', 'Source!A1');
    const task = {
      ...program(),
      evidence: { kind: 'cell-sources' as const, sources: [sourceVersion(source)] },
    };
    expect(
      (
        await validateDurableSources(
          task,
          bridge(async () => source),
        )
      ).status,
    ).toBe('current');
    for (const current of [
      await cell('updated', 'sheet-source', 'Source!A1'),
      await cell('source', 'other-sheet', 'Source!A1'),
    ]) {
      expect(
        (
          await validateDurableSources(
            task,
            bridge(async () => current),
          )
        ).status,
      ).toBe('changed');
    }
    expect(
      (
        await validateDurableSources(
          task,
          bridge(async () => ({ ...source, values: [['different']] })),
        )
      ).status,
    ).toBe('unsupported');
    task.evidence.sources[0]!.locator = 'Source!A:A';
    const capture = vi.fn(async () => source);
    expect((await validateDurableSources(task, bridge(capture))).status).toBe('unsupported');
    expect(capture).not.toHaveBeenCalled();
  });
  it('checks effect-specific lineage in addition to its current output', async () => {
    const { input } = await cellEvidence();
    const source = await cell('source', 'sheet-source', 'Source!A1');
    input.preconditions!.push(sourceVersion(source));
    const evidence = deriveDurableInspection(input)!;
    const output = await cell();
    const captures = vi.fn(async (selector: string) =>
      selector === 'Output!A1' ? output : cell('changed', 'sheet-source', 'Source!A1'),
    );
    expect(
      (await inspectDurableOutput(evidence, receipt('write-cells', output.hash), bridge(captures)))
        .status,
    ).toBe('changed');
    expect(captures).toHaveBeenCalledTimes(2);
  });
  it('requires disjoint actual output footprints including anchored grids and native object aliases', async () => {
    const first = program('grid Output!A1 = "a\\tb\\nc\\td"\nfinish when=verified');
    first.stages.push({
      id: 'second',
      dependsOn: ['write'],
      program: 'set output!B2 "collision"\nfinish when=verified',
    });
    expect((await validateDurableSources(first, bridge())).status).toBe('unsupported');
    first.stages[1]!.program = 'set Output!D2 "separate"\nfinish when=verified';
    expect((await validateDurableSources(first, bridge())).status).toBe('current');
    const word = {
      ...program('/fill-content-control contentControlId=1 text="first"\nfinish when=verified'),
      surface: 'word' as const,
    };
    word.stages.push({
      id: 'second',
      dependsOn: ['write'],
      program: '/fill-content-control contentControlId=01 text="second"\nfinish when=verified',
    });
    expect(
      (await validateDurableSources(word, { surface: 'word', documentKey: () => 'saved' })).status,
    ).toBe('unsupported');
    const ppt = {
      ...program(
        '/set-shape-text slideId=slide-1 shapeId=shape-1 text="first"\nfinish when=verified',
      ),
      surface: 'powerpoint' as const,
    };
    ppt.stages.push({
      id: 'second',
      dependsOn: ['write'],
      program:
        '/set-shape-text slideId=slide-1 shapeId=shape-1 text="second"\nfinish when=verified',
    });
    expect(
      (await validateDurableSources(ppt, { surface: 'powerpoint', documentKey: () => 'saved' }))
        .status,
    ).toBe('unsupported');
  });
  it('refuses declared source/output overlap before capture or the first effect', async () => {
    const source = await cell('source', 'sheet-output', 'Output!A1');
    const task = {
      ...program(),
      evidence: { kind: 'cell-sources' as const, sources: [sourceVersion(source)] },
    };
    const capture = vi.fn(async () => source);
    const assessment = await validateDurableSources(task, bridge(capture));
    expect(assessment.status).toBe('unsupported');
    expect(assessment.reason).toContain('own declared source');
    expect(capture).not.toHaveBeenCalled();
  });
});
