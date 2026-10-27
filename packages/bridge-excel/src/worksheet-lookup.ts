import { BridgeRefusal } from '@ge/runtime';

/**
 * Worksheet lookup through the null-object path.
 *
 * `worksheets.getItem(name)` on a missing sheet surfaces only at the next `sync()` as an opaque
 * RichApi `ItemNotFound`, which the dispatcher then has to treat as an unknown outcome. The
 * `getItemOrNullObject` form lets the bridge name the missing sheet in one clear refusal BEFORE
 * any write is queued, so a typo in a sheet-qualified address is a corrective the model can act
 * on. The existence probe costs one bounded sync and is read-only, which is also why a cell in
 * edit mode met here is a certain non-write (the same reasoning as the protected-sheet probe).
 */

/** The message a missing worksheet produces; shared so tests and callers match one wording. */
export function missingWorksheetMessage(name: string): string {
  return `Worksheet "${name}" was not found in this workbook.`;
}

/** Minimal structural slice of the worksheet collection the lookup needs (fakes implement it). */
interface WorksheetLookupContext {
  workbook: {
    worksheets: {
      getItemOrNullObject(name: string): { isNullObject: boolean; load(props: string): unknown };
    };
  };
  sync(): Promise<unknown>;
}

type LookedUp<C extends WorksheetLookupContext> = ReturnType<
  C['workbook']['worksheets']['getItemOrNullObject']
>;

/**
 * Resolve one or more worksheets by exact name in a single sync. Raises a {@link BridgeRefusal}
 * (`not_found`) for the first name that does not exist, or (`edit_mode`) when the probe itself is
 * refused because a cell is being edited; both are certain non-writes the dispatcher reports as
 * `host_refused`. Call it FIRST in a batch, before any write statement is queued.
 */
export async function resolveWorksheets<C extends WorksheetLookupContext>(
  ctx: C,
  names: readonly string[],
): Promise<Array<LookedUp<C>>> {
  const sheets = names.map((name) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(name);
    sheet.load('isNullObject');
    return sheet;
  });
  try {
    await ctx.sync();
  } catch (error) {
    if ((error as { code?: unknown } | null)?.code === 'InvalidOperationInCellEditMode')
      throw new BridgeRefusal('edit_mode', 'A cell is being edited in Excel.');
    throw error;
  }
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i]!.isNullObject)
      throw new BridgeRefusal('not_found', missingWorksheetMessage(names[i]!), names[i]!);
  }
  return sheets as Array<LookedUp<C>>;
}

/** Two-sheet convenience (source + destination) over {@link resolveWorksheets}: one sync. */
export async function resolveWorksheetPair<C extends WorksheetLookupContext>(
  ctx: C,
  first: string,
  second: string,
): Promise<[LookedUp<C>, LookedUp<C>]> {
  const [a, b] = await resolveWorksheets(ctx, [first, second]);
  return [a!, b!];
}

/** Single-sheet convenience over {@link resolveWorksheets}. */
export async function resolveWorksheet<C extends WorksheetLookupContext>(
  ctx: C,
  name: string,
): Promise<LookedUp<C>> {
  const [sheet] = await resolveWorksheets(ctx, [name]);
  return sheet!;
}
