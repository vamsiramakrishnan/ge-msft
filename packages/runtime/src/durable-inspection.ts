import {
  ActuationRequestSchema,
  CellSnapshotSchema,
  DurableInspectionEvidenceSchema,
  DurableCommandProgramSchema,
  assessActuationResult,
  asChangeId,
  boundedExcelRange,
  contentHash,
  gridForRequest,
  isProgramCommand,
  isProgramVerifiedFinish,
  parseProgramBlock,
  rangeForGrid,
  type ActuationRequest,
  type ActuationResult,
  type CellSnapshot,
  type ContextRef,
  type DurableCommandProgram,
  type DurableInspectionEvidence,
  type SourceVersion,
} from '@ge/contracts';
import type { DocBridge } from './bridge.js';
import { compileCommand, isCompileError } from './command-protocol.js';
import { boundedNativeRead, type NativeReadOptions } from './bounded-native-read.js';

export interface DurableInspectionResult {
  status: 'current' | 'changed' | 'unsupported';
  reason?: string;
}
const unsupported = (reason: string): DurableInspectionResult => ({
  status: 'unsupported',
  reason,
});
const changed = (reason: string): DurableInspectionResult => ({ status: 'changed', reason });
const current = (): DurableInspectionResult => ({ status: 'current' });
const hash = /^sha256:[a-f0-9]{64}$/;
const bytes = (input: unknown): number =>
  new TextEncoder().encode(JSON.stringify(input)).byteLength;
const MAX_CAPTURE_BYTES = 1024 * 1024;
const MAX_TOTAL_BYTES = 2 * MAX_CAPTURE_BYTES;

/** Only explicit, stable native identities can support future current-output checks. */
export function deriveDurableInspection(
  raw: ActuationRequest,
  documentId?: string,
): DurableInspectionEvidence | undefined {
  const parsed = ActuationRequestSchema.safeParse(raw);
  if (!parsed.success) return undefined;
  const request = parsed.data;
  const target = request.params.target;
  const conditions = request.preconditions ?? [];
  if (conditions.length > 16) return undefined;
  if (
    request.surface === 'excel' &&
    request.kind === 'write-cells' &&
    target?.range &&
    boundedExcelRange(target.range)
  ) {
    const destination = conditions.find(
      (source) => source.surface === 'excel' && source.locator === target.range && source.objectId,
    );
    if (!destination?.objectId) return undefined;
    return DurableInspectionEvidenceSchema.parse({
      version: 1,
      kind: 'cells',
      operationKind: request.kind,
      changeId: request.changeId,
      documentId: destination.documentId,
      objectId: destination.objectId,
      locator: destination.locator,
      sources: conditions.filter((source) => source !== destination),
    });
  }
  const textDocumentId = documentId ?? target?.expectedDocumentId;
  if (
    !target ||
    !textDocumentId ||
    !hash.test(textDocumentId) ||
    (target.expectedDocumentId && target.expectedDocumentId !== textDocumentId)
  )
    return undefined;
  if (
    request.surface === 'word' &&
    request.kind === 'fill-content-control' &&
    target.contentControlId
  )
    return DurableInspectionEvidenceSchema.parse({
      version: 1,
      kind: 'text',
      surface: 'word',
      operationKind: request.kind,
      changeId: request.changeId,
      documentId: textDocumentId,
      objectId: target.contentControlId,
      sources: conditions,
    });
  if (
    request.surface === 'powerpoint' &&
    request.kind === 'set-shape-text' &&
    target.slideId &&
    target.shapeId
  )
    return DurableInspectionEvidenceSchema.parse({
      version: 1,
      kind: 'text',
      surface: 'powerpoint',
      operationKind: request.kind,
      changeId: request.changeId,
      documentId: textDocumentId,
      objectId: target.shapeId,
      parentObjectId: target.slideId,
      sources: conditions,
    });
  return undefined;
}

type ReadBridge = Pick<DocBridge, 'surface' | 'documentKey' | 'captureCells' | 'captureTextTarget'>;
function sameCells(snapshot: CellSnapshot, source: SourceVersion): boolean {
  return (
    snapshot.surface === source.surface &&
    snapshot.documentId === source.documentId &&
    snapshot.locator === source.locator &&
    snapshot.objectId === source.objectId
  );
}
async function readCells(
  selector: string,
  bridge: ReadBridge,
  options: NativeReadOptions,
): Promise<CellSnapshot> {
  const snapshot = CellSnapshotSchema.parse(
    await boundedNativeRead(() => bridge.captureCells!(selector), options),
  );
  const range = boundedExcelRange(snapshot.locator);
  if (
    !range ||
    snapshot.values.length !== range.rows ||
    snapshot.values[0]!.length !== range.columns
  )
    throw new Error('Native capture did not match its bounded rectangle.');
  if (bytes(snapshot) > MAX_CAPTURE_BYTES)
    throw new Error('Native capture exceeds the inspection byte budget.');
  const actualHash = await contentHash({
    surface: snapshot.surface,
    documentId: snapshot.documentId,
    locator: snapshot.locator,
    objectId: snapshot.objectId,
    values: snapshot.values,
    formulas: snapshot.formulas ?? [],
  });
  if (actualHash !== snapshot.hash)
    throw new Error('Native capture hash does not match its contents.');
  return snapshot;
}
async function checkSources(
  sources: readonly SourceVersion[],
  bridge: ReadBridge,
  budget: { bytes: number },
  options: NativeReadOptions,
): Promise<DurableInspectionResult> {
  if (!sources.length) return current();
  if (sources.length > 16 || !bridge.captureCells || bridge.surface !== 'excel')
    return unsupported('The saved sources need a native read port this host does not provide.');
  for (const source of sources) {
    if (source.surface !== 'excel' || !source.objectId || !boundedExcelRange(source.locator))
      return unsupported('A saved source lacks a bounded native worksheet identity.');
    const snapshot = await readCells(source.locator, bridge, options);
    const size = bytes(snapshot);
    budget.bytes += size;
    if (size > MAX_CAPTURE_BYTES || budget.bytes > MAX_TOTAL_BYTES)
      return unsupported('Native source inspection exceeds the bounded capture budget.');
    if (!sameCells(snapshot, source) || snapshot.hash !== source.hash)
      return changed('A declared source no longer matches its saved native version.');
  }
  return current();
}

/** Literal declaration is an explicit author assertion; expressions or host reads cannot acquire it. */
type BoundedRange = NonNullable<ReturnType<typeof boundedExcelRange>>;
function rangesOverlap(left: BoundedRange, right: BoundedRange): boolean {
  return (
    left.sheet.toUpperCase() === right.sheet.toUpperCase() &&
    left.firstRow < right.firstRow + right.rows &&
    right.firstRow < left.firstRow + left.rows &&
    left.firstColumn < right.firstColumn + right.columns &&
    right.firstColumn < left.firstColumn + left.columns
  );
}
function validateLiteralProgram(program: DurableCommandProgram): DurableInspectionResult {
  const cellOutputs: BoundedRange[] = [];
  const sourceRanges =
    program.evidence?.kind === 'cell-sources'
      ? program.evidence.sources.map((source) => boundedExcelRange(source.locator))
      : [];
  const objectOutputs = new Set<string>();
  for (const stage of program.stages) {
    const parsed = parseProgramBlock('```cmd\n' + stage.program + '\n```');
    let effects = 0;
    let completed = false;
    for (const entry of parsed.entries) {
      if (isProgramVerifiedFinish(entry)) {
        if (effects !== 1 || completed)
          return unsupported('A saved stage must write before its final completion control.');
        completed = true;
        continue;
      }
      if (!isProgramCommand(entry))
        return unsupported(
          'Literal tasks cannot depend on expressions, analysis bindings or macros.',
        );
      if (entry.verb === 'done') {
        if (effects !== 1 || completed)
          return unsupported('A saved stage must write before its final completion control.');
        completed = true;
        continue;
      }
      if (completed)
        return unsupported('Saved stages cannot place commands after a completion control.');
      if (
        ['valueExpr', 'textExpr', 'bodyExpr', 'bulletsExpr'].some((key) =>
          Object.hasOwn(entry, key),
        )
      )
        return unsupported('Literal tasks cannot resolve write arguments from expressions.');
      const compiled = compileCommand(entry, {
        surface: program.surface,
        mintChangeId: () => asChangeId('literal-inspection'),
      });
      if (isCompileError(compiled) || compiled.kind !== 'write')
        return unsupported(
          'Literal tasks contain only exact write commands and completion controls.',
        );
      if (++effects !== 1) return unsupported('Each literal task stage needs exactly one write.');
      const request = compiled.request;
      if (!['write-cells', 'fill-content-control', 'set-shape-text'].includes(request.kind))
        return unsupported(
          'This literal operation has no supported native saved-output inspector.',
        );
      if (request.kind === 'write-cells') {
        if (!request.params.target?.range || !boundedExcelRange(request.params.target.range))
          return unsupported(
            'Literal cell writes need an explicit bounded sheet-qualified rectangle.',
          );
        const grid = gridForRequest(request);
        if (!grid.length || !grid[0]?.length || grid.some((row) => row.length !== grid[0]!.length))
          return unsupported('Literal cell writes need a nonempty rectangular value grid.');
        const footprint = boundedExcelRange(
          rangeForGrid(request.params.target.range, grid.length, grid[0].length),
        );
        if (!footprint)
          return unsupported('Literal cell output exceeds its bounded native rectangle.');
        if (sourceRanges.some((source) => source && rangesOverlap(source, footprint)))
          return unsupported(
            'Automatic inspection cannot overwrite its own declared source ranges.',
          );
        if (cellOutputs.some((prior) => rangesOverlap(prior, footprint)))
          return unsupported(
            'Saved task stages need disjoint output ranges for automatic inspection.',
          );
        cellOutputs.push(footprint);
        if (
          request.params.cellFormulas?.some((row) => row.some(Boolean)) ||
          request.params.cells?.some((row) => row.some((value) => value.startsWith('=')))
        )
          return unsupported(
            'A formula depends on native values and cannot declare literal-only evidence.',
          );
      }
      if (
        (request.kind === 'fill-content-control' && !request.params.target?.contentControlId) ||
        (request.kind === 'set-shape-text' &&
          (!request.params.target?.slideId || !request.params.target.shapeId))
      )
        return unsupported('Literal text writes need exact native object IDs.');
      if (request.kind !== 'write-cells') {
        const target = request.params.target!;
        const wordId = target.contentControlId;
        const canonicalId =
          wordId && wordId.trim() && Number.isFinite(Number(wordId))
            ? String(Number(wordId))
            : wordId;
        const key =
          request.kind === 'fill-content-control'
            ? `word:${canonicalId}`
            : `powerpoint:${target.slideId}:${target.shapeId}`;
        if (objectOutputs.has(key))
          return unsupported(
            'Saved task stages need distinct native text objects for automatic inspection.',
          );
        objectOutputs.add(key);
      }
    }
    if (effects !== 1) return unsupported('Each literal task stage needs exactly one write.');
  }
  return current();
}

/** Validates declared lineage; it never invents source coverage for a legacy or model-produced task. */
export async function validateDurableSources(
  raw: DurableCommandProgram,
  bridge: ReadBridge,
  options: NativeReadOptions = {},
): Promise<DurableInspectionResult> {
  const program = DurableCommandProgramSchema.parse(raw);
  const evidence = program.evidence;
  if (!evidence) return unsupported('This saved task has no explicit evidence declaration.');
  if (program.surface !== bridge.surface)
    return changed('Open the task in its original Office application.');
  const key = bridge.documentKey?.();
  if (!key) return unsupported('Save the Office document before inspecting its task.');
  try {
    const structure = validateLiteralProgram(program);
    const result =
      structure.status !== 'current'
        ? structure
        : evidence.kind === 'literal'
          ? structure
          : await checkSources(
              evidence.sources,
              bridge,
              { bytes: 0 },
              { deadline: Date.now() + 10000, ...options },
            );
    if (bridge.documentKey?.() !== key)
      return changed('The Office document changed during source inspection.');
    return result;
  } catch {
    return unsupported('Native source inspection could not establish currentness.');
  }
}

/** Read-only currentness proof from the actual correlated native postcondition hash. */
export async function inspectDurableOutput(
  raw: DurableInspectionEvidence | undefined,
  receipt: ActuationResult | undefined,
  bridge: ReadBridge,
  readOptions: NativeReadOptions = {},
): Promise<DurableInspectionResult> {
  const parsed = DurableInspectionEvidenceSchema.safeParse(raw);
  if (!parsed.success || !receipt)
    return unsupported('This checkpoint has no supported native inspection evidence.');
  const evidence = parsed.data;
  if (
    receipt.changeId !== evidence.changeId ||
    receipt.kind !== evidence.operationKind ||
    assessActuationResult(receipt) !== 'verified' ||
    receipt.recoveryPending ||
    receipt.trackingModeRestoreFailed ||
    !hash.test(receipt.verification?.afterHash ?? '')
  )
    return unsupported('A correlated verified native postcondition hash is required.');
  const key = bridge.documentKey?.();
  if (!key) return unsupported('Save the Office document before inspecting its output.');
  const budget = { bytes: 0 };
  const options = { deadline: Date.now() + 10000, ...readOptions };
  try {
    if (evidence.kind === 'cells') {
      if (
        bridge.surface !== 'excel' ||
        !bridge.captureCells ||
        !boundedExcelRange(evidence.locator)
      )
        return unsupported('The saved output needs a bounded native cell read port.');
      const snapshot = await readCells(evidence.locator, bridge, options);
      const size = bytes(snapshot);
      budget.bytes += size;
      if (size > MAX_CAPTURE_BYTES)
        return unsupported('Native output inspection exceeds the bounded capture budget.');
      if (
        snapshot.surface !== 'excel' ||
        snapshot.documentId !== evidence.documentId ||
        snapshot.objectId !== evidence.objectId ||
        snapshot.locator !== evidence.locator ||
        snapshot.hash !== receipt.verification!.afterHash
      )
        return changed('The native cell output changed after its verified checkpoint.');
    } else {
      if (bridge.surface !== evidence.surface || !bridge.captureTextTarget)
        return unsupported('The saved output needs a native text-object read port.');
      const ref: ContextRef =
        evidence.surface === 'word'
          ? {
              id: `word:cc:${evidence.objectId}`,
              surface: 'word',
              kind: 'paragraph',
              title: 'Saved task output',
              hostRef: { type: 'word.contentControl', contentControlId: evidence.objectId },
            }
          : {
              id: `pp:shape:${evidence.parentObjectId}:${evidence.objectId}`,
              surface: 'powerpoint',
              kind: 'shape',
              title: 'Saved task output',
              hostRef: {
                type: 'powerpoint.shape',
                slideId: evidence.parentObjectId ?? '',
                shapeId: evidence.objectId,
              },
            };
      const capture = await boundedNativeRead(() => bridge.captureTextTarget!(ref), options);
      if (new TextEncoder().encode(capture.text).byteLength > 64 * 1024)
        return unsupported('Native text output exceeds the inspection budget.');
      budget.bytes += new TextEncoder().encode(capture.text).byteLength;
      if (
        capture.target.surface !== evidence.surface ||
        capture.target.documentId !== evidence.documentId ||
        capture.target.objectId !== evidence.objectId ||
        capture.target.parentObjectId !== evidence.parentObjectId ||
        capture.target.revision !== receipt.verification!.afterHash ||
        (await contentHash({ text: capture.text })) !== receipt.verification!.afterHash
      )
        return changed('The native text output changed after its verified checkpoint.');
    }
    const sources = await checkSources(evidence.sources, bridge, budget, options);
    if (bridge.documentKey?.() !== key)
      return changed('The Office document changed during output inspection.');
    return sources;
  } catch {
    return unsupported('Native output inspection could not establish currentness.');
  }
}
