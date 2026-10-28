import { z } from 'zod';
import { EFFECT_VERBS } from './expr-grammar.js';

/**
 * C1 richer algebra — the bounded `each` loop:
 *
 *   each $row in $table:
 *     set Summary!B2 = $row.amount
 *     comment $row.cell "flagged"
 *   end
 *
 * The block is grouped by `parseProgramBlock` exactly like a `def … end` recipe, and carries its
 * body lines VERBATIM. The runtime (`packages/runtime/src/each-expansion.ts`) substitutes each
 * `$row.<column>` with the cell text of one row of the bound table and expands the body into
 * CONCRETE effect lines BEFORE dry-run — so the approval card shows every write that will land,
 * never a loop. The expansion is bounded (rows and effects) and refused as a whole when any line
 * would not be an effect.
 *
 * This module owns the SHAPE (header syntax, the body rules the Python preflight mirrors); the
 * runtime owns the expansion.
 */

/* ───────────────────────────── AST ───────────────────────────── */

/** A grouped `each $row in $table: … end` block. Names exclude the `$`. */
export interface ParsedEachBlock {
  kind: 'each-block';
  rowVar: string;
  tableVar: string;
  /** Verbatim body lines (trimmed; blanks/comments removed); each starts with an effect verb. */
  body: string[];
}

export const ParsedEachBlockSchema: z.ZodType<ParsedEachBlock> = z.object({
  kind: z.literal('each-block'),
  rowVar: z.string(),
  tableVar: z.string(),
  body: z.array(z.string()),
});

export function isParsedEachBlock(node: { kind?: string }): node is ParsedEachBlock {
  return node.kind === 'each-block';
}

/* ───────────────────────────── header ───────────────────────────── */

const EACH_HEADER_RE = /^each\s+\$([A-Za-z_][A-Za-z0-9_]*)\s+in\s+\$([A-Za-z_][A-Za-z0-9_]*)\s*:$/i;

export const EACH_HEADER_USAGE = 'usage: each $row in $table:  …effect lines…  end';

/** Is this line an attempt at an `each` header (strictly parsed by {@link parseEachHeader})? */
export function isEachHeader(line: string): boolean {
  return /^each(?:\s|$)/i.test(line.trim());
}

/** Parse `each $row in $table:` → `{ rowVar, tableVar }` or a corrective error. */
export function parseEachHeader(
  line: string,
): { rowVar: string; tableVar: string } | { error: string } {
  const m = EACH_HEADER_RE.exec(line.trim());
  if (!m) return { error: `malformed each header — ${EACH_HEADER_USAGE}` };
  const rowVar = m[1]!;
  const tableVar = m[2]!;
  if (rowVar === tableVar) {
    return { error: `each needs distinct names for the row and the table — ${EACH_HEADER_USAGE}` };
  }
  return { rowVar, tableVar };
}

/* ───────────────────────────── body rules ───────────────────────────── */

/** A `$name.column` row reference (group 1 = the variable, group 2 = the column). */
export const ROW_REF_RE = /\$([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)/g;

/**
 * Validate one body line against the structural rules the TS and Python parsers share:
 *   • it must start with an EFFECT verb (reads, `let`, pipelines, `def`, `each` are refused — a
 *     loop body is a template for writes, nothing else);
 *   • every `$x.col` reference must name the loop's row variable;
 *   • the row variable may only be used as `$row.<column>` (a bare `$row` has no value).
 * Column existence is checked at expansion time (the table is only known then).
 */
export function eachBodyLineError(rowVar: string, line: string): string | undefined {
  const sp = line.search(/\s/);
  const head = (sp === -1 ? line : line.slice(0, sp)).toLowerCase();
  if (!EFFECT_VERBS.has(head)) {
    return `each body may only contain effect commands (${[...EFFECT_VERBS].join(', ')}) — got "${head}"`;
  }
  for (const m of line.matchAll(ROW_REF_RE)) {
    if (m[1] !== rowVar) {
      return `unknown row variable $${m[1]} — the loop variable is $${rowVar}`;
    }
  }
  const bare = new RegExp(`\\$${rowVar}(?![A-Za-z0-9_.])`);
  if (bare.test(line)) {
    return `use $${rowVar}.<column> to reference a cell — a bare $${rowVar} has no value`;
  }
  return undefined;
}
