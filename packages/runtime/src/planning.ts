import {
  approvalClassOf,
  effectIdempotencyKey,
  isReversibleKind,
  type ActuationRequest,
  type ApprovalClass,
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

/** A sheet name as Excel means it: `'O''Brien'` is O'Brien. */
function sheetName(raw: string): string {
  const name = raw.trim();
  const quoted = /^'(.*)'$/.exec(name);
  return quoted ? quoted[1]!.replace(/''/g, "'") : name;
}

function parseRange(ref: string): ParsedRange | undefined {
  const m = RANGE_RE.exec(ref.trim());
  if (!m) return undefined;
  const c1 = colToNum(m[2]!);
  const r1 = Number(m[3]);
  const c2 = m[4] ? colToNum(m[4]) : c1;
  const r2 = m[5] ? Number(m[5]) : r1;
  return {
    sheet: sheetName(m[1] ?? ''),
    c1: Math.min(c1, c2),
    r1: Math.min(r1, r2),
    c2: Math.max(c1, c2),
    r2: Math.max(r1, r2),
  };
}

/** Split a comma-separated area list on commas outside quoted sheet names. */
function splitAreas(address: string): string[] {
  const out: string[] = [];
  let current = '';
  let quoted = false;
  for (const ch of address) {
    if (ch === "'") quoted = !quoted;
    if (ch === ',' && !quoted) {
      out.push(current.trim());
      current = '';
    } else current += ch;
  }
  out.push(current.trim());
  return out.filter(Boolean);
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

type Grid = readonly (readonly unknown[])[];
interface CellCapture {
  locator: string;
  values: Grid;
  formulas?: Grid;
}

/** The value and formula of one cell of a capture, comparable across captures. */
const cellKey = (grid: CellCapture, r: number, c: number): string =>
  JSON.stringify([grid.values[r]?.[c], grid.formulas?.[r]?.[c] ?? '']);

/**
 * True when `after` differs from `before` (two captures of `locator`) and every differing cell holds
 * exactly what a capture in `landed` holds for that cell: the cells an earlier effect of the same
 * batch produced. Fails closed: an unparseable locator, another sheet or a different shape is never
 * covered, and neither is a cell whose content differs from what the batch produced.
 */
export function changesCoveredBy(
  before: CellCapture,
  after: CellCapture,
  landed: readonly CellCapture[],
): boolean {
  const origin = parseRange(before.locator);
  if (!origin?.sheet || before.values.length !== after.values.length) return false;
  const areas = landed.flatMap((capture) => {
    const area = parseRange(capture.locator);
    return area && area.sheet.toLowerCase() === origin.sheet.toLowerCase()
      ? [{ area, capture }]
      : [];
  });
  let changed = 0;
  for (let r = 0; r < before.values.length; r++) {
    const width = before.values[r]!.length;
    if (after.values[r]!.length !== width) return false;
    for (let c = 0; c < width; c++) {
      if (cellKey(before, r, c) === cellKey(after, r, c)) continue;
      changed++;
      const row = origin.r1 + r;
      const col = origin.c1 + c;
      const covered = areas.some(
        ({ area, capture }) =>
          area.r1 <= row &&
          row <= area.r2 &&
          area.c1 <= col &&
          col <= area.c2 &&
          cellKey(capture, row - area.r1, col - area.c1) === cellKey(after, r, c),
      );
      if (!covered) return false;
    }
  }
  return changed > 0;
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
    case 'sort-range':
      return { reads: [], writes: p.sortRange?.range ? [range(p.sortRange.range)] : [] };
    case 'create-table':
      return {
        reads: p.table?.range ? [range(p.table.range)] : [],
        writes: [obj(`table:${p.table?.name ?? p.table?.range ?? req.changeId}`)],
      };
    case 'insert-chart':
      // A multi-area source (`Sheet2!C1:C11,Sheet2!G1:G11`) depends on each of its areas.
      return {
        reads: p.chart?.sourceRange
          ? splitAreas(p.chart.sourceRange).map((area) => range(area))
          : [],
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
    case 'insert-slide':
      // Appends a slide, which becomes the deck's LAST slide: a later `slide=last` write reads it, so
      // a failed/uncertain insert blocks that write instead of it landing on the user's own last slide.
      return { reads: [], writes: [{ kind: 'estate', id: req.kind }, obj(LAST_SLIDE)] };
    case 'add-shape':
    case 'set-shape-text':
    case 'format-shape':
      return {
        reads: isLastSlideRef(p.target?.slideId) ? [obj(LAST_SLIDE)] : [],
        writes: [{ kind: 'estate', id: req.kind }],
      };
    case 'add-table-slide':
      // `slide=new` appends a slide (it becomes the last slide, like insert-slide); otherwise it
      // writes onto an existing slide, depending on a preceding insert when that slide is `last`.
      return isNewSlideRef(p.target?.slideId)
        ? { reads: [], writes: [{ kind: 'estate', id: req.kind }, obj(LAST_SLIDE)] }
        : {
            reads: isLastSlideRef(p.target?.slideId) ? [obj(LAST_SLIDE)] : [],
            writes: [{ kind: 'estate', id: req.kind }],
          };
    default: {
      // Anchored content writes (tracked-change/insert-*/fill-content-control) + estate kinds.
      const anchor = p.target?.matchText ?? p.target?.contentControlId;
      if (anchor) return { reads: [], writes: [{ kind: 'anchor', id: anchor }] };
      // A range-addressed kind without a dedicated case (filter, validation, …) writes its range, so
      // a later write into the same cells is ordered after it instead of racing it.
      if (p.target?.range) return { reads: [], writes: [range(p.target.range)] };
      return { reads: [], writes: [{ kind: 'estate', id: req.kind }] };
    }
  }
}

/** The PowerPoint deck's last slide, as a plan resource (see the `insert-slide` case above). */
const LAST_SLIDE = 'pp:slide:last';

/** A PowerPoint slide reference meaning "the deck's last slide" (resolved by the bridge). */
function isLastSlideRef(slideId: string | undefined): boolean {
  return slideId !== undefined && /^(?:pp:)?(?:slide:)?last$/i.test(slideId.trim());
}

/** `/add-table-slide slide=new`: the bridge creates the slide. */
function isNewSlideRef(slideId: string | undefined): boolean {
  return slideId !== undefined && /^(?:pp:)?(?:slide:)?new$/i.test(slideId.trim());
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
      // Stable across retries: the intent hash, not the per-attempt changeId (B4).
      idempotencyKey: effectIdempotencyKey(request),
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

// ───────────────────────────── scheduling (B1) ─────────────────────────────

/** Severity order of approval authorities — the order phases are taken in. */
export const APPROVAL_CLASS_ORDER: readonly ApprovalClass[] = [
  'in-document',
  'external',
  'estate',
  'irreversible',
];

/**
 * One batch of effects with no dependency between them (Kahn level). Every node in a level depends
 * only on nodes in EARLIER levels, so a level's effects may run in any order — or, once a bridge
 * supports it, be dispatched together. Node order inside a level is source order.
 */
export interface ScheduledLevel {
  level: number;
  ids: string[];
}

/**
 * Partition the DAG into dependency levels (Kahn's algorithm, stable by source order). Level 0 holds
 * the roots; a node sits at `1 + max(level of its prerequisites)`. The output lists every node once
 * (the DAG is acyclic by construction: `dependsOn` only names earlier ids). The executor still
 * dispatches ONE effect at a time — the task journal is single-writer and Office hosts serialise
 * sync batches — but level order lets independent effects proceed past a failed sibling and gives
 * the preview an honest "these N can land together" grouping instead of a flat list.
 */
export function scheduleEffectLevels(nodes: readonly EffectPlanNode[]): ScheduledLevel[] {
  const level = new Map<string, number>();
  for (const node of nodes) {
    let depth = 0;
    for (const dep of node.dependsOn) {
      const upstream = level.get(dep);
      if (upstream !== undefined) depth = Math.max(depth, upstream + 1);
    }
    level.set(node.id, depth);
  }
  const buckets = new Map<number, string[]>();
  for (const node of nodes) {
    const depth = level.get(node.id)!;
    const ids = buckets.get(depth) ?? [];
    ids.push(node.id);
    buckets.set(depth, ids);
  }
  return [...buckets.keys()]
    .sort((a, b) => a - b)
    .map((depth) => ({ level: depth, ids: buckets.get(depth)! }));
}

/** One approval phase: the effects of ONE approval class, taken as one decision. */
export interface ApprovalPhase {
  approvalClass: ApprovalClass;
  ids: string[];
}

/**
 * Split a plan into per-class approval phases, severity-ordered (ADR-0008 §break-boundaries): the
 * user decides each authority separately — the in-document edits first, then the external sends,
 * then estate writes, then anything irreversible. A phase is only sound when nothing in it depends
 * on an effect of a LATER phase (an in-document write that reads a slide an irreversible layout
 * change produces would otherwise run before its prerequisite was ever approved). When that
 * happens the plan collapses to ONE mixed phase — still carrying every class, so the approval
 * surface knows the decision bundles authorities and must say so. A single-class plan is one phase.
 */
export function approvalPhases(nodes: readonly EffectPlanNode[]): ApprovalPhase[] {
  const present = APPROVAL_CLASS_ORDER.filter((cls) =>
    nodes.some((node) => node.approvalClass === cls),
  );
  if (present.length <= 1)
    return present.map((approvalClass) => ({
      approvalClass,
      ids: nodes.map((node) => node.id),
    }));
  const rank = new Map(present.map((cls, i) => [cls, i]));
  const classOf = new Map(nodes.map((node) => [node.id, node.approvalClass]));
  const crossesBack = nodes.some((node) =>
    node.dependsOn.some((dep) => {
      const upstream = classOf.get(dep);
      return upstream !== undefined && rank.get(upstream)! > rank.get(node.approvalClass)!;
    }),
  );
  if (crossesBack)
    return [{ approvalClass: present[present.length - 1]!, ids: nodes.map((node) => node.id) }];
  return present.map((approvalClass) => ({
    approvalClass,
    ids: nodes.filter((node) => node.approvalClass === approvalClass).map((node) => node.id),
  }));
}

/**
 * The dispatch order of a plan: phase by phase (severity-ordered, so nothing irreversible runs
 * before the cheaper phases landed), and level by level inside a phase. Returns node ids.
 */
export function executionOrder(nodes: readonly EffectPlanNode[]): string[] {
  const byId = new Map(nodes.map((node) => [node.id, node]));
  const out: string[] = [];
  for (const phase of approvalPhases(nodes)) {
    const members = phase.ids.map((id) => byId.get(id)!);
    for (const level of scheduleEffectLevels(members)) out.push(...level.ids);
  }
  return out;
}
