/** Revision-local evidence links from captured native formulas. These are never write locators. */
export interface FormulaSheet {
  id: string;
  name: string;
}
export interface FormulaDependencies {
  relatedIds: string[];
  complete: boolean;
  limitations: string[];
}
export interface FormulaAlias {
  name: string;
  address: string;
  kind: 'named-range' | 'table';
}

const MAX_FORMULA_LENGTH = 16 * 1024;
const MAX_LINKS = 100;
const LIMITATION =
  'Formula dependencies cover literal A1 and bounded native aliases; dynamic, external, unresolved named/structured, whole-row/column and 3D references require a fresh explicit read. Links do not grant write authority.';

/** Absolute grid buckets make cross-sheet links stable without loading another sheet's values. */
export function excelRowUnit(sheetId: string, row: number, column: number): string {
  return `sheet:${sheetId}/row:${row}/col:${Math.floor((column - 1) / 32) * 32 + 1}`;
}

/**
 * Deliberately not an Excel evaluator. Only literal native A1 references are resolved. Strings
 * and unsupported address mechanisms cannot manufacture edges. Oversized ranges stop at the
 * explicit link budget; consumers must preserve complete=false when using the graph.
 */
export function formulaDependencies(
  formula: string,
  currentSheet: FormulaSheet,
  sheets: readonly FormulaSheet[],
  maxLinks = MAX_LINKS,
  aliases: readonly FormulaAlias[] = [],
): FormulaDependencies {
  if (!Number.isSafeInteger(maxLinks) || maxLinks < 0 || maxLinks > MAX_LINKS)
    throw new Error('Invalid dependency link budget.');
  if (aliases.length > 256) throw new Error('Invalid dependency alias budget.');
  const result: FormulaDependencies = { relatedIds: [], complete: true, limitations: [] };
  const incomplete = () => {
    result.complete = false;
    if (!result.limitations.length) result.limitations.push(LIMITATION);
  };
  if (!formula.startsWith('=') || formula.length > MAX_FORMULA_LENGTH) {
    incomplete();
    return result;
  }
  // Blank strings before lexing; quoted sheet names use single quotes and remain address data.
  let source = formula.replace(/"(?:[^"]|"")*"/g, (match) => ' '.repeat(match.length));
  let expandedLength = source.length;
  const expand = (alias: FormulaAlias | undefined, original: string) => {
    if (!alias || alias.address.length > 4096) return original;
    const replacement = `(${alias.address.replace(/^=/, '')})`;
    const nextLength = expandedLength + replacement.length - original.length;
    if (nextLength > 64 * 1024) {
      incomplete();
      return original;
    }
    expandedLength = nextLength;
    return replacement;
  };
  // Native aliases carry addresses, never values or authority. Structured references use the
  // entire native table footprint conservatively; header/totals/row selectors cannot under-read.
  const lookup = (name: string, kind: FormulaAlias['kind']) => {
    const unquoted = name.replace(
      /^'((?:[^']|'')+)'!/,
      (_match, sheet: string) => `${sheet.replace(/''/g, "'")}!`,
    );
    const local = `${currentSheet.name}!${unquoted}`.toLocaleLowerCase();
    return (
      aliases.find((alias) => alias.kind === kind && alias.name.toLocaleLowerCase() === local) ??
      aliases.find(
        (alias) =>
          alias.kind === kind && alias.name.toLocaleLowerCase() === unquoted.toLocaleLowerCase(),
      )
    );
  };
  source = source.replace(
    /([\p{L}_\\][\p{L}\p{N}_.\\]*)\[((?:[^[\]]|\[[^\]]*\])*)\]/gu,
    (match, name: string) => {
      const alias = lookup(name, 'table');
      return expand(alias, match);
    },
  );
  source = source.replace(
    /(?:(?:'(?:[^']|'')+'|[\p{L}_\\][\p{L}\p{N}_.\\]*)!)?[\p{L}_\\][\p{L}\p{N}_.\\]*/gu,
    (match, offset: number, whole: string) => {
      if (/^\s*\(/.test(whole.slice(offset + match.length))) return match;
      const alias = lookup(match, 'named-range');
      return expand(alias, match);
    },
  );
  if (/\b(?:INDIRECT|OFFSET)\s*\(/i.test(source) || /[[\]#]/.test(source)) incomplete();
  const seen = new Set<string>();
  const replacements: Array<[number, number]> = [];
  const references =
    /(?:(?:'((?:[^']|'')+)'|([\p{L}_\\][\p{L}\p{N}_.\\]*))!)?(\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6})(?::(\$?[A-Z]{1,3}\$?[1-9][0-9]{0,6}))?/giu;
  for (const match of source.matchAll(references)) {
    const start = match.index!;
    const end = start + match[0].length;
    if (
      /[\p{L}\p{N}_.[\]:]/u.test(source[start - 1] ?? '') ||
      /[\p{L}\p{N}_.[\](]/u.test(source[end] ?? '')
    )
      continue;
    const sheetName = match[1]?.replace(/''/g, "'") ?? match[2];
    const sheet =
      sheetName === undefined
        ? currentSheet
        : sheets.find((item) => item.name.toLocaleLowerCase() === sheetName.toLocaleLowerCase());
    const from = cell(match[3]!);
    const to = cell(match[4] ?? match[3]!);
    replacements.push([start, end]);
    if (!sheet || !from || !to || from.row > to.row || from.column > to.column) {
      incomplete();
      continue;
    }
    // At most 100 links, even for a million-row range. Never expand the full address first.
    for (let row = from.row; row <= to.row; row++) {
      for (let column = from.column; column <= to.column; ) {
        const id = excelRowUnit(sheet.id, row, column);
        if (!seen.has(id)) {
          if (seen.size >= maxLinks) {
            incomplete();
            return result;
          }
          seen.add(id);
          result.relatedIds.push(id);
        }
        column = (Math.floor((column - 1) / 32) + 1) * 32 + 1;
      }
    }
  }
  for (const [start, end] of replacements.reverse())
    source = source.slice(0, start) + ' '.repeat(end - start) + source.slice(end);
  // Remaining identifiers are functions/constants or unresolved names/structured/3D references.
  source = source.replace(/[\p{L}_\\][\p{L}\p{N}_.\\]*\s*(?=\()/gu, ' ');
  source = source.replace(/\b(?:TRUE|FALSE|E[+-]?\d+)\b/gi, ' ');
  if (/[\p{L}_\\]/u.test(source)) incomplete();
  return result;
}

function cell(value: string): { row: number; column: number } | undefined {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})$/i.exec(value);
  if (!match) return;
  const column = [...match[1]!.toUpperCase()].reduce(
    (sum, char) => sum * 26 + char.charCodeAt(0) - 64,
    0,
  );
  const row = Number(match[2]);
  if (row > 1048576 || column > 16384) return;
  return { row, column };
}
