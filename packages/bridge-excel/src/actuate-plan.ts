import type { ActuationRequest } from '@ge/contracts';

/**
 * Pure translation of an actuation into a host plan — testable without Office.js. A
 * `write-cells` is located by an explicit `target.range` (e.g. "Sheet1!A1:B3"); the grid to
 * write is `params.cells`. The bridge parses the address and writes `range.values` at
 * apply-time.
 */
export interface WriteCellsPlan {
  address?: string;
  values: string[][];
}

export function planWriteCells(req: ActuationRequest): WriteCellsPlan {
  const p = req.params;
  return {
    ...(p.target?.range ? { address: p.target.range } : {}),
    values: p.cells ?? [],
  };
}
