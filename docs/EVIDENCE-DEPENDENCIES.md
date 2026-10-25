# Revision-bound dependency navigation

Excel paged capture now derives read-only evidence links from native `Range.formulas`, worksheet
IDs and bounded native named-range/table metadata. Literal A1 references resolve to the captured
row buckets that contain the referenced cells, including other worksheets. Structured references
cover the full native table range conservatively. A table header, total or row selector therefore
cannot exclude a potentially relevant table cell.

The adapter counts workbook names, current-worksheet names and tables before loading each
collection, admits at most 128 from each collection, and retains at most 256 aliases and 16 KiB
of alias metadata per tile. Local names shadow workbook names. Unknown local coverage prevents
unqualified global name resolution. It loads addresses, not alias/table values.
These metadata loads share the existing three native synchronization stages. This establishes a
bounded plan, not a claim of reduced native latency. Cells remain limited to 1,024 per tile.

Row IDs use absolute 32-column grid buckets, so an A1 reference can identify another worksheet's
evidence without loading that worksheet's values. A reference outside the captured used range
remains unresolved. These are evidence IDs in a particular revision; they are not native write
locators or evidence-binding authority.

The formula index handles literal A1 ranges, qualified/quoted worksheet names, direct native
range names and structured table references. It is not an Excel formula evaluator. Dynamic,
external, whole-row/column, 3D, missing aliases and oversized references remain explicit gaps.
Formula-looking literal strings never create dependency links. Per-row links are capped at 100;
individual formulas at 16,384 UTF-16 code units and alias expansion at 65,536 code units. A limitation is
retained when the adapter cannot establish complete relation coverage.

## SDK use

`AssistSession.sourceWorkspace` exposes the same immutable evidence workspace used by DocFs.
Indexing is incremental and explicit:

```ts
const workspace = session.sourceWorkspace;
let progress;
do {
  progress = await workspace.indexDependencies(snapshotId, { maxPages: 8, signal });
} while (!progress.complete && !progress.limited);

const dependents = await workspace.query(snapshotId, {
  kind: 'expand', id: capturedUnitId, direction: 'dependents',
});
// Follow nextCursor when the query page is incomplete, as with ordinary evidence navigation.
```

The index reads at most eight authenticated pages per call and reuses pages with identical hashes
across revisions. Source pages are independently validated before entering the reusable cache.
Word/PowerPoint parent links and other adapter-declared relationships use the same index. Existing
PowerPoint slide/shape containment is retained; this does not invent a Word clause dependency or
claim complete semantic dependency extraction.

After indexing two revisions of the same source and scope:

```ts
const impact = workspace.dependencyImpact(previousSnapshotId, currentSnapshotId, {
  maxNodes: 256, maxDepth: 8,
});
if (impact.requiresFullRefresh) {
  // Incomplete, unknown, unresolved or over-budget relationships cannot certify unaffected units.
  // Refresh the complete source before deriving a result that relies on selective invalidation.
}
```

Impact follows old and new reverse links, so removed dependencies and cycles are handled. It
includes changed, added and removed units, plus downstream captured dependents. Parent edges
conservatively invalidate captured children. Direct `dependents` navigation follows formula/related
links; `children` remains the containment query. Returned units and limitations are data.

The index retains at most four revisions, 16,000 units, 64,000 links and 4 MiB of encoded node
metadata per revision. It refuses additional revisions until an accelerator is explicitly released:

```ts
await workspace.releaseDependencyIndex(oldSnapshotId); // Retains the source evidence.
await workspace.release(oldSnapshotId); // Releases source evidence and its accelerator.
```

Limits are visible through `limited`, `omittedUnits`, `unresolvedLinks` and
`incompleteRelationUnits`. Legacy snapshots without the optional `relationsComplete` declaration
have unknown relation coverage. A partial index never certifies selective invalidation.

Dependency impact never replaces the existing complete-source freshness check, target capture,
capability negotiation or approval. An unchanged graph can still refer to stale native data.
No write dispatch, approval, credential, host transaction or automatic rebasing is introduced.

## Verification boundary

The simulator tests verify native ID binding, string/external/dynamic reference handling,
range/alias limits, native metadata counts and three-stage batching. Revision fixtures verify
eight-page indexing, nine unchanged pages reused out of ten, reverse invalidation, cycles,
legacy/partial coverage, corrupt pages, cancellation, cache quotas and competing calls. They also
verify that a source edit is still refused by the independent full-source freshness check.
These are deterministic local measurements; native Office performance and semantic parity remain
separate certification checks.
