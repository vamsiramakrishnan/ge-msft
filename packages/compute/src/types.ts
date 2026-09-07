import type { CellValue, TableArtifact } from '@ge/contracts';
export interface QueryResult {
  columns: string[];
  rows: CellValue[][];
  truncated: boolean;
  durationMs: number;
}
export interface ComputeEngine {
  query(sql: string, tables: readonly TableArtifact[], signal?: AbortSignal): Promise<QueryResult>;
  dispose(): void;
}
