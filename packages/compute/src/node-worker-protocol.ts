import type { TableArtifact } from '@ge/contracts';

/** Internal correlated messages; no source text or errors are logged by the worker. */
export interface NodeComputeJob {
  version: 1;
  id: string;
  sql: string;
  tables: TableArtifact[];
  maxRows: number;
}
