import type { TableArtifact } from '@ge/contracts';

/** Native doubles above this magnitude may already have lost input digits before capture. */
export function assertExactDecimalColumn(table: TableArtifact, index: number): void {
  if (!table.columns[index])
    throw new Error(
      `Column ${index + 1} (index ${index}) does not exist in ${table.title}; the captured range has ${table.columns.length} columns. Expand the range or choose another column.`,
    );
  if (
    table.rows.some(
      (row) =>
        typeof row[index] === 'number' && Math.abs(row[index] as number) > Number.MAX_SAFE_INTEGER,
    )
  )
    throw new Error(
      `Amount column ${index + 1} in ${table.title} contains a number beyond safe numeric precision. Store those amounts as decimal text and capture the range again.`,
    );
}

/** DECIMAL(38,6) admission: invalid precision becomes NULL instead of being silently rounded. */
export function exactDecimalColumnSql(column: string): string {
  if (!/^c\d+$/.test(column))
    throw new Error('An exact decimal expression requires a captured column identifier.');
  const text = `trim(cast(${column} as varchar))`;
  return `case when regexp_full_match(${text}, '[+-]?[0-9]{1,32}([.][0-9]{1,6})?') then try_cast(${text} as decimal(38,6)) else null end`;
}
