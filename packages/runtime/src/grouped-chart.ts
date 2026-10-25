/**
 * `chart line Sheet!A1:H25 x=Month by=Branch value=Complaints` — a trend per group from long data
 * (one row per group per month). The pane builds the summary table itself, from the cells it read,
 * so no total is typed by the model (UAT 2026-10-16: the chart showed 24 raw bars, or numbers
 * the model wrote by hand). Each summary cell is a live SUMPRODUCT over the source columns (exact match).
 */

type Cell = string | number | boolean | null | undefined;

export interface GroupedChartPlan {
  /** Where the summary goes: x labels down the first column, one column per group. */
  range: string;
  cells: string[][];
}

const columnIndex = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const columnLabel = (index: number): string => {
  let s = '';
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

/** `'Branch Performance'!A1:H25` → its parts; undefined for anything else. */
export function parseArea(
  address: string,
):
  | { prefix: string; firstColumn: number; firstRow: number; lastColumn: number; lastRow: number }
  | undefined {
  const bang = address.lastIndexOf('!');
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(address.slice(bang + 1));
  if (!m) return undefined;
  return {
    prefix: bang < 0 ? '' : address.slice(0, bang + 1),
    firstColumn: columnIndex(m[1]!),
    firstRow: Number(m[2]),
    lastColumn: columnIndex(m[3]!),
    lastRow: Number(m[4]),
  };
}

/** The address of the summary block, `gap` columns to the right of the source. */
export function summaryRange(
  source: string,
  groups: number,
  xValues: number,
  offset: number,
): string | undefined {
  const a = parseArea(source);
  if (!a) return undefined;
  const left = a.lastColumn + 2 + offset;
  const right = left + groups;
  if (right > 16384 || a.firstRow + xValues > 1048576) return undefined;
  return `${a.prefix}${columnLabel(left)}${a.firstRow}:${columnLabel(right)}${a.firstRow + xValues}`;
}

/**
 * The summary table for a grouped chart, or an error the model can act on. `values` is the source
 * range as read, header row first. `x`, `by`, `value` name a header (case-insensitive) or a
 * column letter inside the source.
 */
export function groupedChartPlan(
  source: string,
  values: Cell[][],
  props: { x?: string; by?: string; value?: string },
  offset = 0,
): GroupedChartPlan | { error: string } {
  const area = parseArea(source);
  if (!area)
    return { error: `chart ${source}: a grouped chart needs a block range like Sheet!A1:H25` };
  const header = (values[0] ?? []).map((h) => String(h ?? '').trim());
  const find = (name: string | undefined, key: string): number | { error: string } => {
    if (!name)
      return {
        error: `chart needs ${key}=<column> with by= (e.g. x=Month by=Branch value=Complaints)`,
      };
    const i = header.findIndex((h) => h.toLowerCase() === name.trim().toLowerCase());
    if (i >= 0) return i;
    if (/^[A-Za-z]{1,3}$/.test(name)) {
      const j = columnIndex(name) - area.firstColumn;
      if (j >= 0 && j < header.length) return j;
    }
    return {
      error: `chart ${key}=${name}: no such column; the headers are ${header.filter(Boolean).join(', ')}`,
    };
  };
  const xi = find(props.x, 'x');
  const bi = find(props.by, 'by');
  const vi = find(props.value, 'value');
  for (const i of [xi, bi, vi]) if (typeof i !== 'number') return i;
  const [x, by, value] = [xi, bi, vi] as number[];

  const rows = values
    .slice(1)
    .filter((r) => r.some((c) => c !== null && c !== undefined && c !== ''));
  // One label per value as Excel compares them: case-insensitive, exactly as written in the cell.
  const distinct = (i: number): string[] => {
    const seen = new Map<string, string>();
    for (const r of rows) {
      const label = String(r[i] ?? '');
      if (label.trim() && !seen.has(label.toLowerCase())) seen.set(label.toLowerCase(), label);
    }
    return [...seen.values()];
  };
  const xs = distinct(x!);
  const groups = distinct(by!);
  if (xs.length === 0 || groups.length === 0)
    return {
      error: `chart ${source}: no rows to summarise under ${header[x!]} and ${header[by!]}`,
    };
  if (groups.length > 50 || xs.length > 500)
    return {
      error: `chart ${source}: too many groups (${groups.length}) or x values (${xs.length}) to chart`,
    };

  const range = summaryRange(source, groups.length, xs.length, offset);
  const dest = range ? parseArea(range) : undefined;
  if (!range || !dest)
    return { error: `chart ${source}: no room on the sheet for the summary table` };
  const col = (i: number): string => {
    const letter = columnLabel(area.firstColumn + i);
    return `$${letter}$${area.firstRow + 1}:$${letter}$${area.lastRow}`;
  };
  // Labels come from the document: one that starts like a formula is written as text, never run.
  const text = (label: string): string => (/^[=+\-@]/.test(label) ? `'${label}` : label);
  const cells: string[][] = [[text(header[x!] || 'X'), ...groups.map(text)]];
  xs.forEach((label, r) => {
    const row = dest.firstRow + 1 + r;
    cells.push([
      text(label),
      ...groups.map((_, g) => {
        const groupCell = `${columnLabel(dest.firstColumn + 1 + g)}$${dest.firstRow}`;
        const xCell = `$${columnLabel(dest.firstColumn)}${row}`;
        // Exact match: SUMIFS would read `*`, `?`, `>100` in a label as a pattern or a test.
        return `=SUMPRODUCT(--(${col(by!)}=${groupCell}),--(${col(x!)}=${xCell}),${col(value!)})`;
      }),
    ]);
  });
  return { range, cells };
}
