import { ROW_REF_RE, scanQuoted, type ParsedEachBlock } from '@ge/contracts';
import type { Value } from './compose.js';

/**
 * C1 — expansion of a bounded `each $row in $table: … end` block into CONCRETE effect lines.
 *
 * The loop is sugar, not a runtime construct: before dry-run, every body line is instantiated once
 * per row of the bound table with each `$row.<column>` replaced by that row's cell text, and the
 * resulting lines run through the ordinary program path (parse → compile → dry-run → ONE approval
 * → actuate → receipt). The approval card therefore shows every write that will land; nothing is
 * decided after approval.
 *
 * Bounds (fail closed, refused as a whole — never a partial expansion):
 *   • at most {@link MAX_EACH_ROWS} rows;
 *   • at most `limits.maxEffects` expanded lines (the per-turn write cap);
 *   • every `$row.<column>` must name a column of the table.
 *
 * Cell text is HOST CONTENT — data, never instructions. Substitution keeps it data: a cell can
 * never add a line (line breaks are flattened), never close or open a quoted span (quotes and
 * backslashes are escaped inside quotes), and outside quotes it may not carry the characters that
 * would re-shape a command (`|`, `(`, `)`, `"`). Such a cell must be referenced inside a quoted
 * argument, where it is inert.
 */

/** The largest table an `each` may iterate (a loop over more is a job for a filtered table). */
export const MAX_EACH_ROWS = 50;

export interface EachLimits {
  maxRows: number;
  /** The per-turn write cap: the expansion may not produce more effect lines than this. */
  maxEffects: number;
}

export type EachExpansion =
  | { ok: true; lines: string[]; rows: number }
  | { ok: false; error: string };

/** Characters a substituted cell may not carry OUTSIDE a quoted argument. */
const UNSAFE_UNQUOTED = /[|()"\\]/;

export function expandEachBlock(
  block: ParsedEachBlock,
  env: ReadonlyMap<string, Value>,
  limits: EachLimits,
): EachExpansion {
  const { rowVar, tableVar, body } = block;
  const table = env.get(tableVar);
  if (!table) {
    return { ok: false, error: `each: unbound variable "$${tableVar}" — bind it with let first` };
  }
  if (table.kind !== 'table') {
    return { ok: false, error: `each: $${tableVar} is a ${table.kind}, not a table` };
  }
  if (table.rows.length > limits.maxRows) {
    return {
      ok: false,
      error: `each: $${tableVar} has ${table.rows.length} rows; the loop limit is ${limits.maxRows} — filter or head the table first`,
    };
  }
  const total = table.rows.length * body.length;
  if (total > limits.maxEffects) {
    return {
      ok: false,
      error: `each: ${table.rows.length} rows × ${body.length} line(s) = ${total} effects exceeds this turn's write cap of ${limits.maxEffects} — narrow the table or split the work`,
    };
  }

  // Resolve every referenced column once, up front — an unknown column refuses the whole block.
  const refs = new Map<string, number>();
  for (const line of body) {
    for (const m of line.matchAll(ROW_REF_RE)) {
      const col = m[2]!;
      if (refs.has(col)) continue;
      const idx = table.columns.indexOf(col);
      if (idx === -1) {
        return {
          ok: false,
          error: `each: $${rowVar}.${col} — $${tableVar} has no column "${col}" (columns: ${table.columns.join(', ')})`,
        };
      }
      refs.set(col, idx);
    }
  }

  const lines: string[] = [];
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r]!;
    for (const line of body) {
      const out = substituteRow(line, rowVar, row, refs, r + 1);
      if ('error' in out) return { ok: false, error: out.error };
      lines.push(out.line);
    }
  }
  return { ok: true, lines, rows: table.rows.length };
}

/**
 * Substitute `$row.<col>` in one body line. The line is walked as quoted / unquoted segments so the
 * escaping rule depends on where the reference sits.
 */
function substituteRow(
  line: string,
  rowVar: string,
  row: string[],
  refs: ReadonlyMap<string, number>,
  rowNo: number,
): { line: string } | { error: string } {
  let out = '';
  let i = 0;
  while (i < line.length) {
    if (line[i] === '"') {
      const q = scanQuoted(line, i);
      if (q) {
        const raw = line.slice(i, q.end);
        out += replaceRefs(raw, rowVar, (col) => escapeQuoted(cell(row, refs, col)));
        i = q.end;
        continue;
      }
    }
    // Unquoted run: up to the next quote (or the end).
    let j = i + 1;
    while (j < line.length && line[j] !== '"') j++;
    const raw = line.slice(i, j);
    let bad: string | undefined;
    out += replaceRefs(raw, rowVar, (col) => {
      const text = flatten(cell(row, refs, col));
      if (UNSAFE_UNQUOTED.test(text) && bad === undefined) {
        bad = `each: $${rowVar}.${col} (row ${rowNo}) contains | ( ) " or \\ — reference it inside a quoted argument`;
      }
      return text;
    });
    if (bad !== undefined) return { error: bad };
    i = j;
  }
  return { line: out };
}

function cell(row: string[], refs: ReadonlyMap<string, number>, col: string): string {
  return row[refs.get(col)!] ?? '';
}

function replaceRefs(text: string, rowVar: string, sub: (col: string) => string): string {
  return text.replace(ROW_REF_RE, (whole: string, name: string, col: string) =>
    name === rowVar ? sub(col) : whole,
  );
}

/** Line breaks can never add a program line. */
function flatten(text: string): string {
  return text.replace(/[\r\n]+/g, ' ');
}

/** Inside a quoted argument a cell stays one string: escape backslashes and quotes, flatten lines. */
function escapeQuoted(text: string): string {
  return flatten(text).replace(/\\/g, '\\\\').replace(/"/g, '\\"');
}
