import {
  DeckTemplateCaptureSchema,
  boundedExcelRange,
  type ActuationKind,
  type DeckTemplateCapture,
  type HostCapabilityProfile,
  type Surface,
} from '@ge/contracts';

export const NATIVE_FIXTURE_VERSION = 1;
export interface NativeFixtureExpectation {
  surface: Surface;
  expectedEffects: readonly ActuationKind[];
  verifiedEffects?: readonly number[];
  reviewOpened?: boolean;
}
const EXPECTATIONS: Readonly<Record<string, NativeFixtureExpectation>> = {
  'word-tracked-edit': {
    surface: 'word',
    expectedEffects: ['insert-text', 'tracked-change'],
    verifiedEffects: [1],
  },
  'word-table-readback': {
    surface: 'word',
    expectedEffects: ['insert-text', 'insert-table'],
    verifiedEffects: [1],
  },
  'excel-grid-readback': {
    surface: 'excel',
    expectedEffects: ['write-cells'],
    verifiedEffects: [0],
  },
  'excel-table-readback': {
    surface: 'excel',
    expectedEffects: ['write-cells', 'create-table'],
    verifiedEffects: [0, 1],
  },
  'excel-pivot-readback': {
    surface: 'excel',
    expectedEffects: ['write-cells', 'insert-pivot'],
    verifiedEffects: [0, 1],
  },
  'excel-filter-readback': {
    surface: 'excel',
    expectedEffects: ['write-cells', 'filter-range'],
    verifiedEffects: [0, 1],
  },
  'excel-validation-readback': {
    surface: 'excel',
    expectedEffects: ['write-cells', 'set-data-validation'],
    verifiedEffects: [0, 1],
  },
  'powerpoint-slide-readback': { surface: 'powerpoint', expectedEffects: ['insert-slide'] },
  'powerpoint-layout-readback': {
    surface: 'powerpoint',
    expectedEffects: ['apply-slide-layout'],
    verifiedEffects: [0],
  },
  'outlook-review-form': {
    surface: 'outlook',
    expectedEffects: ['create-mail'],
    reviewOpened: true,
  },
  'outlook-compose-readback': {
    surface: 'outlook',
    expectedEffects: ['set-subject', 'set-body'],
    verifiedEffects: [0, 1],
  },
};
/** Shared by report collection and execution; imported Boolean checks alone cannot prove a case. */
export function nativeFixtureExpectation(id: string): NativeFixtureExpectation | undefined {
  const expectation = Object.hasOwn(EXPECTATIONS, id) ? EXPECTATIONS[id] : undefined;
  return expectation ? structuredClone(expectation) : undefined;
}
export interface NativeFixture {
  id: string;
  surface: Surface;
  /** Setup is separately dry-run and approved so subsequent operations capture the landed seed. */
  setupProgram?: string;
  program: string;
  expectedKind: ActuationKind;
  expectedEffects: readonly ActuationKind[];
  /** Zero-based receipt indices whose actual native verifier must pass. */
  verifiedEffects?: readonly number[];
  expectedText?: string;
  reviewOpened?: boolean;
  modes: readonly HostCapabilityProfile['mode'][];
  cleanup: string;
}
const registered = new WeakMap<NativeFixture, string>();
const quote = (value: string): string =>
  JSON.stringify(value)
    .replace(/`/g, '\\u0060')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
function seal(
  definition: Omit<
    NativeFixture,
    'expectedKind' | 'expectedEffects' | 'verifiedEffects' | 'reviewOpened'
  >,
): NativeFixture {
  const expected = nativeFixtureExpectation(definition.id);
  if (!expected || expected.surface !== definition.surface)
    throw new Error('Unknown native fixture definition.');
  const fixture: NativeFixture = {
    ...definition,
    ...expected,
    expectedEffects: Object.freeze([...expected.expectedEffects]),
    ...(expected.verifiedEffects
      ? { verifiedEffects: Object.freeze([...expected.verifiedEffects]) }
      : {}),
    modes: Object.freeze([...definition.modes]),
    expectedKind: expected.expectedEffects[expected.expectedEffects.length - 1]!,
  };
  Object.freeze(fixture);
  registered.set(fixture, JSON.stringify(fixture));
  return fixture;
}
/** A canonical fixture ID cannot be relabelled onto an edited or deserialized command program. */
export function isNativeFixture(fixture: NativeFixture): boolean {
  const fingerprint = registered.get(fixture);
  return fingerprint !== undefined && fingerprint === JSON.stringify(fixture);
}
const nonce = (): string => `GE_CERT_${crypto.randomUUID().replace(/-/g, '')}`;

function explicitRange(value: string, rows: number, columns: number): string {
  const range = boundedExcelRange(value);
  if (!range || range.rows !== rows || range.columns !== columns || /[\r\n`]/.test(value))
    throw new Error(`Supply an explicit disposable ${rows}×${columns} sheet-qualified range.`);
  return range.address;
}

/** Run only in a disposable file/draft. Normal approval cards show every fixture mutation. */
export function nativeFixture(surface: Surface, targetRange?: string): NativeFixture {
  const marker = nonce();
  if (surface === 'word')
    return seal({
      id: 'word-tracked-edit',
      surface,
      program: `/insert-text position=end text=${quote(marker + '_OLD')}\nsuggest ${quote(marker + '_OLD')} => ${quote(marker + '_NEW')}`,
      expectedText: marker + '_NEW',
      modes: ['write'],
      cleanup: 'Reject or remove the fixture tracked edits in the disposable document.',
    });
  if (surface === 'excel') {
    if (!targetRange) throw new Error('Supply an explicit disposable workbook target range.');
    const range = explicitRange(targetRange, 2, 2);
    return seal({
      id: 'excel-grid-readback',
      surface,
      program: `grid ${quote(range)} = ${quote('Metric\tValue\n' + marker + '\t12.34')}\nfinish when=verified`,
      modes: ['write'],
      cleanup: 'Use the reviewed Recovery & undo action to restore the fixture range.',
    });
  }
  if (surface === 'powerpoint')
    return seal({
      id: 'powerpoint-slide-readback',
      surface,
      program: `slide ${quote(marker)} "Native text readback"`,
      expectedText: marker,
      modes: ['write'],
      cleanup: 'Remove the fixture slide from the disposable presentation.',
    });
  if (surface === 'outlook')
    return seal({
      id: 'outlook-review-form',
      surface,
      program: `compose ${quote(marker)} "Native certification: review form only."`,
      modes: ['read', 'compose'],
      cleanup: 'Discard the fixture draft without sending it.',
    });
  throw new Error('No native certification fixture is promoted for this surface.');
}

export const NATIVE_OPERATION_FIXTURE_IDS = [
  'word-table-readback',
  'excel-table-readback',
  'excel-pivot-readback',
  'excel-filter-readback',
  'excel-validation-readback',
  'powerpoint-layout-readback',
  'outlook-compose-readback',
] as const;
export type NativeOperationFixtureId = (typeof NATIVE_OPERATION_FIXTURE_IDS)[number];
export type NativeOperationFixtureInput =
  | { id: 'word-table-readback' | 'outlook-compose-readback'; disposable: true }
  | { id: 'excel-table-readback' | 'excel-filter-readback'; disposable: true; sourceRange: string }
  | { id: 'excel-validation-readback'; disposable: true; targetRange: string }
  | { id: 'excel-pivot-readback'; disposable: true; sourceRange: string; destinationRange: string }
  | {
      id: 'powerpoint-layout-readback';
      disposable: true;
      capture: DeckTemplateCapture;
      slideId: string;
      masterId: string;
      layoutId: string;
    };

/** Fixed bounded seeds and explicit disposable destinations; never creates a substitute static object. */
export function nativeOperationFixture(input: NativeOperationFixtureInput): NativeFixture {
  if (input.disposable !== true)
    throw new Error('Confirm a disposable native fixture destination.');
  const marker = nonce();
  if (input.id === 'word-table-readback')
    return seal({
      id: input.id,
      surface: 'word',
      setupProgram: `/insert-text position=end text=${quote(marker)}`,
      program: `/insert-table match=${quote(marker)} rows=${quote('Metric\tValue\nFixture\t' + marker)}\nfinish when=verified`,
      modes: ['write'],
      cleanup:
        'Remove the fixture table and marker in the disposable document. No durable table undo is claimed.',
    });
  if (input.id === 'outlook-compose-readback')
    return seal({
      id: input.id,
      surface: 'outlook',
      program: `/set-subject subject=${quote(marker)}\n/set-body text=${quote(marker + '_BODY')}\nfinish when=verified`,
      modes: ['compose'],
      cleanup:
        'Discard the saved disposable draft. This fixture never sends mail or certifies on-send event delivery.',
    });
  if (input.id === 'powerpoint-layout-readback') {
    const capture = DeckTemplateCaptureSchema.parse(input.capture);
    const layout = capture.masters
      .find((master) => master.id === input.masterId)
      ?.layouts.find((value) => value.id === input.layoutId);
    if (!capture.complete || !layout || !/^[\s\S]{1,256}$/.test(input.slideId))
      throw new Error('Use a complete native layout capture and explicit disposable slide.');
    return seal({
      id: input.id,
      surface: 'powerpoint',
      program: `/apply-slide-layout slideId=${quote(input.slideId)} masterId=${quote(input.masterId)} layoutId=${quote(input.layoutId)} expectedLayoutHash=${quote(layout.revision)} expectedDocumentId=${quote(capture.documentId)}\nfinish when=verified`,
      modes: ['write'],
      cleanup:
        'Restore the disposable presentation from its saved copy; automatic layout undo is unavailable.',
    });
  }
  if (input.id === 'excel-validation-readback') {
    const range = explicitRange(input.targetRange, 3, 1);
    return seal({
      id: input.id,
      surface: 'excel',
      setupProgram: `grid ${quote(range)} = ${quote('Open\nClosed\nOpen')}\nfinish when=verified`,
      program: `/set-data-validation range=${quote(range)} values="Open,Closed"\nfinish when=verified`,
      modes: ['write'],
      cleanup:
        'Remove the fixture validation rule manually and restore the disposable range; cell undo does not remove validation.',
    });
  }
  if (!('sourceRange' in input)) throw new Error('Unsupported native fixture input.');
  const source = explicitRange(input.sourceRange, 3, 2);
  const setupProgram = `grid ${quote(source)} = ${quote('Region\tAmount\nNorth\t=12.34\nSouth\t=7.66')}\nfinish when=verified`;
  if (input.id === 'excel-table-readback')
    return seal({
      id: input.id,
      surface: 'excel',
      setupProgram,
      program: `table ${quote(source)} headers=true\nfinish when=verified`,
      modes: ['write'],
      cleanup:
        'Remove the minted native table, then restore the seeded cells in the disposable workbook.',
    });
  if (input.id === 'excel-filter-readback')
    return seal({
      id: input.id,
      surface: 'excel',
      setupProgram,
      program: `/filter-range range=${quote(source)} column=Region criterion1=North\nfinish when=verified`,
      modes: ['write'],
      cleanup:
        'Remove the native filter manually and restore the disposable cells; automatic filter undo is unavailable.',
    });
  if (!('destinationRange' in input)) throw new Error('Supply the native pivot destination.');
  const destination = boundedExcelRange(input.destinationRange);
  if (
    !destination ||
    destination.sheet === boundedExcelRange(source)!.sheet ||
    destination.rows !== 6 ||
    destination.columns !== 3 ||
    /[\r\n`]/.test(input.destinationRange)
  )
    throw new Error('Reserve an empty disposable 6×3 rectangle on a different worksheet.');
  return seal({
    id: input.id,
    surface: 'excel',
    setupProgram,
    program: `/insert-pivot sourceRange=${quote(source)} destinationRange=${quote(destination.address)} rowFields=Region valueFields=Amount\nfinish when=verified`,
    modes: ['write'],
    cleanup:
      'Remove the minted native PivotTable and restore the source cells. The reserved destination must be empty before running.',
  });
}
