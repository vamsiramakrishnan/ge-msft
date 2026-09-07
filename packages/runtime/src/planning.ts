import {
  approvalClassOf,
  isReversibleKind,
  type ActuationRequest,
  type EffectPlanNode,
  type ResourceRef,
} from '@ge/contracts';

/**
 * ADR-0008 §7 — compile a sequence of effect requests into a dependency-aware DAG. Dependencies are
 * INFERRED (the model never writes `depends-on`): from overlapping read/write resources, derived
 * ranges, and minted-object references. The result is a saga with bounded compensation — on a node's
 * failure its dependents are skipped; independent effects may still run.
 *
 * This module is pure: it computes the plan shape (reads/writes/deps/approval/failure), it does NOT
 * actuate. Execution + gating happen downstream in the bridge/gate.
 */

// ───────────────────────────── A1 range algebra ─────────────────────────────

interface ParsedRange {
  sheet: string;
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

const RANGE_RE = /^(?:([^!]+)!)?([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/;

function colToNum(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function parseRange(ref: string): ParsedRange | undefined {
  const m = RANGE_RE.exec(ref.trim());
  if (!m) return undefined;
  const c1 = colToNum(m[2]!);
  const r1 = Number(m[3]);
  const c2 = m[4] ? colToNum(m[4]) : c1;
  const r2 = m[5] ? Number(m[5]) : r1;
  return {
    sheet: (m[1] ?? '').replace(/'/g, '').trim(),
    c1: Math.min(c1, c2),
    r1: Math.min(r1, r2),
    c2: Math.max(c1, c2),
    r2: Math.max(r1, r2),
  };
}

/** Expand a grid origin cell by an R×C body (the spill region) so dependents can overlap it. */
function expandGrid(origin: ParsedRange, rows: number, cols: number): ParsedRange {
  if (rows <= 0 || cols <= 0) return origin;
  return { ...origin, c2: origin.c1 + cols - 1, r2: origin.r1 + rows - 1 };
}

function rangesOverlap(a: ParsedRange, b: ParsedRange): boolean {
  if (a.sheet && b.sheet && a.sheet.toLowerCase() !== b.sheet.toLowerCase()) return false;
  return !(a.c2 < b.c1 || b.c2 < a.c1 || a.r2 < b.r1 || b.r2 < a.r1);
}

/** Two resource refs touch the same resource: ranges by geometry, everything else by exact id. */
function refsTouch(a: ResourceRef, b: ResourceRef): boolean {
  if (a.kind === 'range' && b.kind === 'range') {
    const pa = parseRange(a.id);
    const pb = parseRange(b.id);
    return pa !== undefined && pb !== undefined && rangesOverlap(pa, pb);
  }
  return a.kind === b.kind && a.id === b.id;
}

// ───────────────────────────── resource extraction ─────────────────────────────

/**
 * The resources an effect READS and WRITES, by kind. Modeled so the canonical spill→table/chart
 * family infers correctly: `spill`/`set`/`format` WRITE a range; `table`/`chart`/`cf` READ their
 * range (so they depend on the spill that wrote it) and WRITE a minted object (so they don't
 * spuriously depend on each other). Anchored/estate effects use opaque refs.
 */
export function effectResources(req: ActuationRequest): {
  reads: ResourceRef[];
  writes: ResourceRef[];
} {
  const p = req.params;
  const range = (id: string): ResourceRef => ({ kind: 'range', id });
  const obj = (id: string): ResourceRef => ({ kind: 'object', id });

  switch (req.kind) {
    case 'write-cells': {
      // A spilled grid (cells present) writes the whole region, not just the origin cell.
      const origin = p.target?.range;
      if (!origin) return { reads: [], writes: [] };
      const parsed = parseRange(origin);
      const rows = (p.cellValues ?? p.cells)?.length ?? 0;
      const cols = (p.cellValues ?? p.cells)?.[0]?.length ?? 0;
      const region = parsed && rows > 0 ? expandGrid(parsed, rows, cols) : parsed;
      const id = region ? rangeToA1(region) : origin;
      return { reads: [], writes: [range(id)] };
    }
    case 'format-cells':
      return { reads: [], writes: p.target?.range ? [range(p.target.range)] : [] };
    case 'create-table':
      return {
        reads: p.table?.range ? [range(p.table.range)] : [],
        writes: [obj(`table:${p.table?.name ?? p.table?.range ?? req.changeId}`)],
      };
    case 'insert-chart':
      return {
        reads: p.chart?.sourceRange ? [range(p.chart.sourceRange)] : [],
        writes: [obj(`chart:${p.chart?.sourceRange ?? req.changeId}`)],
      };
    case 'format-conditional':
      return {
        reads: p.conditional?.range ? [range(p.conditional.range)] : [],
        writes: [obj(`cf:${p.conditional?.range ?? req.changeId}`)],
      };
    case 'add-comment':
    case 'comment-reply':
      return {
        reads: [],
        writes: [{ kind: 'comment', id: p.target?.commentId ?? p.target?.range ?? req.changeId }],
      };
    case 'reply-mail':
    case 'create-mail':
    case 'post-message':
    case 'post-card':
      return { reads: [], writes: [{ kind: 'draft', id: req.kind }] };
    default: {
      // Anchored content writes (tracked-change/insert-*/fill-content-control) + estate kinds.
      const anchor = p.target?.matchText ?? p.target?.contentControlId;
      if (anchor) return { reads: [], writes: [{ kind: 'anchor', id: anchor }] };
      return { reads: [], writes: [{ kind: 'estate', id: req.kind }] };
    }
  }
}

function rangeToA1(r: ParsedRange): string {
  const col = (n: number): string => {
    let s = '';
    let x = n;
    while (x > 0) {
      const m = (x - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      x = Math.floor((x - 1) / 26);
    }
    return s;
  };
  const a = `${col(r.c1)}${r.r1}`;
  const b = `${col(r.c2)}${r.r2}`;
  const cells = a === b ? a : `${a}:${b}`;
  return r.sheet ? `${r.sheet}!${cells}` : cells;
}

// ───────────────────────────── DAG construction ─────────────────────────────

/**
 * Build the effect DAG from an ordered list of requests. Effect j dependsOn an earlier effect i when
 * j READS something i WROTE (data dependency / derived range) or j WRITES where i WROTE (ordering).
 * Dependencies are inferred — never authored. Returns nodes in source order, each with its resources,
 * approval class, reversibility, idempotency key, and failure policy.
 */
export function analyseEffectDependencies(requests: readonly ActuationRequest[]): EffectPlanNode[] {
  const resources = requests.map(effectResources);
  return requests.map((request, j) => {
    const me = resources[j]!;
    const dependsOn: string[] = [];
    for (let i = 0; i < j; i++) {
      const up = resources[i]!;
      const dataDep = me.reads.some((r) => up.writes.some((w) => refsTouch(r, w)));
      const orderDep = me.writes.some((w) => up.writes.some((w2) => refsTouch(w, w2)));
      if (dataDep || orderDep) dependsOn.push(`e${i + 1}`);
    }
    return {
      kind: 'effect' as const,
      id: `e${j + 1}`,
      dependsOn,
      request,
      reads: me.reads,
      writes: me.writes,
      approvalClass: approvalClassOf(request.kind),
      reversible: isReversibleKind(request.kind),
      idempotencyKey: request.changeId,
      failurePolicy: 'stop-dependents' as const,
    };
  });
}

/**
 * Given a plan and a FAILED node id, the set of dependent nodes that must be skipped
 * (`prerequisite_failed`) — the transitive closure of "depends on the failed node". Independent
 * nodes are not in the set (they may still run if policy permits). The failed node itself is excluded.
 */
export function propagateFailure(
  nodes: readonly EffectPlanNode[],
  failedId: string,
): { skipped: string[]; reason: 'prerequisite_failed' } {
  const failed = new Set([failedId]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const n of nodes) {
      if (!failed.has(n.id) && n.dependsOn.some((d) => failed.has(d))) {
        failed.add(n.id);
        changed = true;
      }
    }
  }
  failed.delete(failedId);
  return { skipped: [...failed], reason: 'prerequisite_failed' };
}
