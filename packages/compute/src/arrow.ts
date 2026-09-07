import { DataType, Table, tableToIPC, vectorFromArray, Utf8, Float64, Bool } from 'apache-arrow';
import type { CellValue, TableArtifact } from '@ge/contracts';
export const ENGINE_SETTINGS =
  "SET autoinstall_known_extensions=false; SET autoload_known_extensions=false; SET enable_external_access=false; SET memory_limit='128MB'; SET threads=1; SET lock_configuration=true";
export function artifactToArrow(artifact: TableArtifact): Table {
  const columns = Object.fromEntries(
    artifact.columns.map((c, i) => {
      const values = artifact.rows.map((r) => r[i] ?? null);
      const vector =
        c.type === 'number'
          ? vectorFromArray(
              values.map((v) => (v === null || v === '' ? null : Number(v))),
              new Float64(),
            )
          : c.type === 'boolean'
            ? vectorFromArray(
                values.map((v) => (v === null || v === '' ? null : Boolean(v))),
                new Bool(),
              )
            : vectorFromArray(
                values.map((v) => (v === null ? null : String(v))),
                new Utf8(),
              );
      return [c.name, vector];
    }),
  );
  return new Table(columns);
}
export function arrowRows(result: Table, limit: number): CellValue[][] {
  if (result.schema.fields.length > 256) throw new Error('Query returned too many columns.');
  const rows: CellValue[][] = [];
  let outputBytes = 0;
  for (let row = 0; row < Math.min(result.numRows, limit); row++) {
    const values = result.schema.fields.map((field, col): CellValue => {
      const value: unknown = result.getChildAt(col)!.get(row);
      if (value == null) return null;
      if (DataType.isDecimal(field.type)) {
        const raw = String(value);
        const scale = field.type.scale;
        const negative = raw.startsWith('-');
        const digits = raw.replace(/^-/, '').padStart(scale + 1, '0');
        return `${negative ? '-' : ''}${scale ? `${digits.slice(0, -scale)}.${digits.slice(-scale)}` : digits}`;
      }
      if (typeof value === 'bigint') return value.toString();
      if (typeof value === 'string' || typeof value === 'boolean') return value;
      if (typeof value === 'number' && Number.isFinite(value)) return value;
      throw new Error('Query returned an unsupported value type. Cast nested values to VARCHAR.');
    });
    outputBytes += new TextEncoder().encode(JSON.stringify(values)).byteLength;
    if (outputBytes > 8 * 1024 * 1024 || values.length * (row + 1) > 100_000)
      throw new Error('Query output exceeds the result budget. Select fewer columns or rows.');
    rows.push(values);
  }
  return rows;
}

export function artifactToIPC(artifact: TableArtifact): Uint8Array {
  return tableToIPC(artifactToArrow(artifact), 'stream');
}
