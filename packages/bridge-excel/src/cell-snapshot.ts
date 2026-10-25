import {
  CellGridSchema,
  MAX_SNAPSHOT_CELLS,
  contentHash,
  makeCellSnapshot,
  type CellSnapshot,
} from '@ge/contracts';

export async function excelDocumentId(): Promise<string> {
  const url = typeof Office !== 'undefined' ? Office.context.document?.url : undefined;
  // Unsaved documents intentionally cannot share a recovery identity across pane lifetimes.
  return url ? contentHash({ host: 'excel', url }) : unsavedId;
}
const unsavedId = `unsaved:${crypto.randomUUID()}`;
export async function snapshotRange(
  range: Excel.Range,
  ctx: Excel.RequestContext,
  documentId: string,
): Promise<CellSnapshot> {
  return (await snapshotRanges([range], ctx, documentId))[0]!;
}

export async function snapshotRanges(
  ranges: readonly Excel.Range[],
  ctx: Excel.RequestContext,
  documentId: string,
): Promise<CellSnapshot[]> {
  if (!ranges.length || ranges.length > 16)
    throw new Error('Capture between one and sixteen explicit ranges.');
  for (const range of ranges) {
    range.load('address,rowCount,columnCount,isNullObject');
    range.worksheet.load('id');
  }
  await ctx.sync();
  for (const range of ranges) {
    if (
      range.isNullObject ||
      range.rowCount < 1 ||
      range.columnCount < 1 ||
      range.rowCount * range.columnCount > MAX_SNAPSHOT_CELLS ||
      range.columnCount > 256
    )
      throw new Error('Select a range of at most 100,000 cells and 256 columns.');
  }
  const output: CellSnapshot[] = [];
  let bytes = 0;
  for (let start = 0; start < ranges.length; ) {
    let end = start;
    let cells = 0;
    while (end < ranges.length) {
      const range = ranges[end]!;
      const count = range.rowCount * range.columnCount;
      if (end > start && cells + count > MAX_SNAPSHOT_CELLS) break;
      cells += count;
      range.load('values,formulas');
      end++;
    }
    await ctx.sync();
    for (const range of ranges.slice(start, end)) {
      const snapshot = await snapshotLoadedRange(range, ctx, documentId);
      bytes += new TextEncoder().encode(JSON.stringify(snapshot)).byteLength;
      if (bytes > 32 * 1024 * 1024)
        throw new Error('Aggregate source snapshots exceed the 32 MiB batch budget.');
      output.push(snapshot);
    }
    start = end;
  }
  return output;
}

async function snapshotLoadedRange(
  range: Excel.Range,
  ctx: Excel.RequestContext,
  documentId: string,
): Promise<CellSnapshot> {
  const worksheet = range.worksheet;
  const values = CellGridSchema.parse(range.values);
  if (new TextEncoder().encode(JSON.stringify(values)).byteLength > 16 * 1024 * 1024)
    throw new Error('Source range exceeds the 16 MiB snapshot budget.');
  const formulas = values.map((row, r) =>
    row.map((_v, c) =>
      typeof range.formulas?.[r]?.[c] === 'string' && String(range.formulas[r]![c]).startsWith('=')
        ? String(range.formulas[r]![c])
        : '',
    ),
  );
  // Range.formulas also returns literal values. A literal '=text' must never become an undo formula.
  if (values.some((row, r) => row.some((v, c) => formulas[r]?.[c] && v === formulas[r]?.[c]))) {
    if (typeof range.getSpecialCellsOrNullObject !== 'function')
      throw new Error(
        'This host cannot distinguish formula-looking text from formulas. Update Office before analyzing this range.',
      );
    const formulaAreas = range.getSpecialCellsOrNullObject('Formulas' as Excel.SpecialCellType);
    formulaAreas.load('isNullObject');
    await ctx.sync();
    const positions = new Set<string>();
    if (!formulaAreas.isNullObject) {
      formulaAreas.areas.load('items/address');
      await ctx.sync();
      const origin = addressSpan(range.address);
      for (const area of formulaAreas.areas.items) {
        const box = addressSpan(area.address);
        for (let r = Math.max(box.r1, origin.r1); r <= Math.min(box.r2, origin.r2); r++)
          for (let c = Math.max(box.c1, origin.c1); c <= Math.min(box.c2, origin.c2); c++)
            positions.add(`${r - origin.r1}:${c - origin.c1}`);
      }
    }
    for (let r = 0; r < formulas.length; r++)
      for (let c = 0; c < formulas[r]!.length; c++)
        if (!positions.has(`${r}:${c}`)) formulas[r]![c] = '';
  }
  if (new TextEncoder().encode(JSON.stringify({ values, formulas })).byteLength > 16 * 1024 * 1024)
    throw new Error('Values and formulas exceed the snapshot budget.');
  return makeCellSnapshot({
    surface: 'excel',
    documentId,
    locator: range.address,
    objectId: worksheet.id,
    values,
    formulas,
  });
}

interface SettingsPort {
  get(key: string): unknown;
  set(key: string, value: unknown): void;
  saveAsync(callback: (result: { status: unknown }) => void): void;
  refreshAsync?(callback: (result: { status: unknown }) => void): void;
}
function settings(): SettingsPort {
  const value = typeof Office !== 'undefined' ? Office.context.document?.settings : undefined;
  if (!value) throw new Error('This host cannot persist recovery records.');
  return value as unknown as SettingsPort;
}
const settingsQueue = new Map<string, Promise<void>>();
export const excelRecoveryStorage = {
  async load(owner: string): Promise<unknown> {
    const key = await storageKey(owner);
    const port = settings();
    if (port.refreshAsync)
      await hostCheckpoint(
        (done) => port.refreshAsync!(done),
        'Recovery history could not be refreshed.',
      );
    const raw = port.get(key);
    if (raw === undefined || raw === null) return [];
    if (typeof raw !== 'string' || raw.length > 2 * 1024 * 1024)
      throw new Error('Recovery history is invalid or too large.');
    return JSON.parse(raw);
  },
  async save(owner: string, value: unknown): Promise<void> {
    const key = await storageKey(owner);
    const payload = JSON.stringify(value);
    if (new TextEncoder().encode(payload).byteLength > 2 * 1024 * 1024)
      throw new Error('Recovery history is full. Resolve or discard older records before writing.');
    const previous = settingsQueue.get(key) ?? Promise.resolve();
    const current = previous
      .catch(() => {})
      .then(async () => {
        const port = settings();
        port.set(key, payload);
        await hostCheckpoint(
          (done) => port.saveAsync(done),
          'Recovery checkpoint could not be saved.',
        );
      });
    settingsQueue.set(key, current);
    try {
      await current;
    } finally {
      if (settingsQueue.get(key) === current) settingsQueue.delete(key);
    }
  },
};
async function storageKey(owner: string): Promise<string> {
  return `ge:recovery:${(await contentHash(owner)).slice(7, 31)}`;
}

function addressSpan(address: string): { r1: number; c1: number; r2: number; c2: number } {
  const match = /([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/.exec(address.replace(/\$/g, ''));
  if (!match) throw new Error('Host returned an invalid formula range.');
  const col = (text: string): number =>
    [...text].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  return {
    r1: Number(match[2]),
    c1: col(match[1]!),
    r2: Number(match[4] ?? match[2]),
    c2: col(match[3] ?? match[1]!),
  };
}

function hostCheckpoint(
  start: (done: (result: { status: unknown }) => void) => void,
  message: string,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${message} Host timed out.`)), 10000);
    try {
      start((result) => {
        clearTimeout(timer);
        if (String(result.status).toLowerCase() === 'succeeded') resolve();
        else reject(new Error(message));
      });
    } catch (error) {
      clearTimeout(timer);
      reject(error);
    }
  });
}
