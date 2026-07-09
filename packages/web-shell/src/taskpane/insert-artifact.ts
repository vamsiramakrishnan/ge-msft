import type { Surface } from '@ge/contracts';

export type InsertableArtifact =
  | {
      kind: 'markdown-table';
      title?: string;
      headers: string[];
      rows: string[][];
    }
  | {
      kind: 'code-block';
      title?: string;
      code: string;
    };

export interface InsertArtifactOptions {
  excelRange?: string;
}

export type InsertArtifactProgram =
  | { ok: true; program: string; label: string }
  | { ok: false; reason: string };

export function buildInsertArtifactProgram(
  surface: Surface,
  artifact: InsertableArtifact,
  opts: InsertArtifactOptions = {},
): InsertArtifactProgram {
  const title =
    artifact.title ?? (artifact.kind === 'markdown-table' ? 'Inserted table' : 'Inserted content');
  const body = artifactToPlainText(artifact);

  switch (surface) {
    case 'excel':
      if (!opts.excelRange) {
        return { ok: false, reason: 'Select a destination range first.' };
      }
      if (artifact.kind === 'markdown-table') {
        return {
          ok: true,
          label: `Insert table into ${opts.excelRange}`,
          program: `grid ${opts.excelRange} = ${quoteGridBody(tableToTsv(artifact))}`,
        };
      }
      return {
        ok: true,
        label: `Insert text into ${opts.excelRange}`,
        program: `set ${opts.excelRange} ${quoteScalar(body)}`,
      };
    case 'word':
      return {
        ok: true,
        label: 'Insert at selection',
        program: `/insert-text text="${propSafeText(body)}"`,
      };
    case 'powerpoint':
      return {
        ok: true,
        label: 'Insert as slide',
        program: `slide ${quoteScalar(title)} ${artifactToSlideBullets(artifact)
          .map(quoteScalar)
          .join(' ')}`,
      };
    case 'onenote':
      return {
        ok: true,
        label: 'Insert as page',
        program: `page ${quoteScalar(title)} ${quoteScalar(body)}`,
      };
    case 'outlook':
      return {
        ok: true,
        label: 'Open draft',
        program: `compose ${quoteScalar(title)} ${quoteScalar(body)}`,
      };
    case 'teams':
      return {
        ok: true,
        label: 'Stage post',
        program: `post ${quoteScalar(body)}`,
      };
  }
}

export function tableToTsv(table: Extract<InsertableArtifact, { kind: 'markdown-table' }>): string {
  return [table.headers, ...table.rows].map((row) => row.map(cleanCell).join('\t')).join('\n');
}

function artifactToPlainText(artifact: InsertableArtifact): string {
  if (artifact.kind === 'code-block') return artifact.code.trim();
  return [artifact.headers, ...artifact.rows]
    .map((row) => row.map(cleanCell).join(' | '))
    .join('\n');
}

function artifactToSlideBullets(artifact: InsertableArtifact): string[] {
  if (artifact.kind === 'code-block') {
    const lines = artifact.code
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    return lines.length > 0 ? lines.slice(0, 8) : ['No content'];
  }
  const rows = artifact.rows.length > 0 ? artifact.rows : [artifact.headers];
  return rows.slice(0, 8).map((row) => row.map(cleanCell).join(' - '));
}

function cleanCell(value: string): string {
  return value.replace(/\*\*/g, '').replace(/`/g, '').replace(/\s+/g, ' ').trim();
}

function quoteGridBody(value: string): string {
  return `"${value
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/\t/g, '\\t')
    .replace(/\r?\n/g, '\\n')}"`;
}

function quoteScalar(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r?\n/g, ' ')}"`;
}

function propSafeText(value: string): string {
  return value.replace(/["\\]/g, "'").replace(/\s+/g, ' ').trim();
}
