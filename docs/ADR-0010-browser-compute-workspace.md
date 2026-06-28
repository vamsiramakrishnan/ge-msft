# ADR-0010 — Browser Compute Workspace (constrained in-browser analysis)

Date: 2026-06-27
Status: Accepted (v1 scoped to Excel + PowerPoint)
Related: ADR-0002 (capability model), ADR-0003 (context construction / Layer-B reads),
ADR-0004 (command protocol → actuation), ADR-0006 (capability closure), ADR-0007 (host-native
write kinds), `docs/research/claude-office-addin-observations.md` (clean-room rationale + the
sanctioned-engine deviation).

## Context

Competitive benchmarking of Anthropic's "Claude in Microsoft Office" (clean-room notes in
`docs/research/claude-office-addin-observations.md`) showed a **constrained in-browser analysis
workspace** — the agent can run tooling over user-approved data snapshots before deciding whether to
call the model, search connectors, or write to the host. Our add-in can ground and emit gated host
edits, but it cannot **run code over data**. Two cases hit a **capability wall** for a fixed verb set:
Excel analysis (joins, group-by, window functions, pivots), and **deck-scale PowerPoint work** —
searching, auditing, and surgically editing hundreds of slides without dumping the whole deck into the
model. Both are solved by the same workspace + tools run over a normalized, *addressable* host
representation that maps back to gated, reversible writes.

This ADR adds that workspace **inside our existing trust model**: client-direct, untrusted host
content, and the unchanged parse → validate → dry-run → preview → approval → policy-gate →
actuation → provenance path for any write.

## Decision

Introduce a surface-agnostic **browser compute workspace**: a memory-only virtual filesystem plus a
sandboxed compute runtime, reached by the model only through **typed Layer-B commands** — never a
general shell.

1. **Compute core (v1):** a no-network **Web Worker** hosting
   - **DuckDB-WASM** as the *single, hardened, read-only* analytical engine — the only code-bearing
     surface; and
   - a fixed registry of **allow-listed safe tools**: `head`, `tail`, `wc`, `rg`, `jq`, `csv-profile`,
     `diff`, `sha256`, and workbook/document outline extractors.
2. **Memory-only VFS:** byte caps, MIME allow-listing, per-file content hashes, explicit lifecycle.
   No durable content/token persistence.
3. **Governance:** compute is read-only and sandboxed, so it **auto-runs** as a Layer-B read; every
   executed command + result preview streams into the **run-steps transcript** (auditable). Any
   write-back to the host goes through the existing gate, unchanged.
4. **v1 scope:** two vertical slices on the shared workspace — **Excel** (tabular analysis + native
   pivot, conditional formatting, tables, number/format) and **PowerPoint** (deck-scale search/audit +
   surgical shape edits) — each writing through the gate. Word surgical breadth and a heavy Python tier
   (Pyodide) are explicitly later phases.

### DuckDB lockdown (non-negotiable)

The DuckDB `query` tool MUST be:
- read-only — no `ATTACH`, no `COPY`-to-disk, no `EXPORT`, no DML against host;
- offline — no `httpfs`, no network of any kind (the Worker has no outbound fetch);
- extension-frozen — no `INSTALL` / `LOAD`;
- filesystem-isolated — no host file access; it operates ONLY over tables the workspace explicitly
  seeds (Arrow/CSV held in memory) from user-approved snapshots/uploads;
- bounded — per-query timeout and memory cap.

The Worker additionally has **no Office.js and no DOM access**. The general-shell / `eval` /
arbitrary-JS / Office-Scripts / PowerShell / process-execution bans from the clean-room doc remain in
force; DuckDB `query` is the sole exception and is sandboxed by construction.

## Architecture

New boundary: `packages/compute` (surface-agnostic, like `runtime`/`web-shell`; no Office.js).

- `Sandbox` (main thread): `seed(name, arrow|csv)`, `query(sql)`, `probe(file, tool, args)` →
  `{ schema, rows, stdout }`, with timeout + memory caps. Spawns and supervises the Worker.
- `worker.ts`: loads DuckDB-WASM (locked per above) + the safe-tool registry over the in-worker VFS.
  No network, no DOM, no Office.js.
- `vfs.ts`: memory-only files with byte caps, MIME allow-list, content hashes, lifecycle/eviction.
- `tools/`: each allow-listed tool as a small pure module with a typed signature.

`runtime` owns orchestration; the surface bridges (`bridge-excel`, `bridge-powerpoint`) own host data
export and host writes. `compute` never touches the host or the network.

## Data flow (analyze → reviewable write)

1. Mid-turn the model emits a typed compute command (below).
2. On first use the active bridge **exports a snapshot** into the workspace via new `DocBridge` hooks:
   `exportTable(selector) → Arrow` (Excel range/sheet) or `exportDeck() → deck files` (PowerPoint
   slides/shapes/notes/theme). `runtime` seeds the result into the VFS.
3. `runtime` runs the command in the Worker; the result re-enters the turn as an **ephemeral context
   block** (like `<doc_state>`, never resident) and appends to the **run-steps transcript**. Auto-run.
4. The model reasons over results → answers in chat, or emits **write commands** that materialise the
   result. Writes traverse the unchanged dry-run → plan-approval → provenance → inverse path.

## Host representation (the addressing linchpin)

The 100x is not the tools — it is exporting the host model **once** into a clean, normalized, *id-keyed*
representation that the tools can slice and that **round-trips to gated writes**. Every exported record
carries the same stable id the write verbs consume, so "find it with a tool" connects directly to "fix
it with a reversible, provenanced write."

- **Excel — `exportTable(selector) → Arrow`**: range/sheet as a columnar table keyed by address; feeds
  DuckDB joins/group-by/window/pivots and `csv-profile` data-quality checks. Writes map back by address.
- **PowerPoint — `exportDeck() → deck files`** into the VFS, keyed by the stable slide/shape ids the
  bridge already uses (`shapeContextRef`, `set-shape-text`, `revealContext`):
  - `slides/NNN.json` — `shapes[{id,type,placeholder,x,y,w,h,z,font,size,bold,color,runs[]}]`, notes, layout
  - `slides/NNN.txt` — flattened text (for `rg`/`wc`)
  - `theme.json` — brand palette + font scheme; `manifest.json` — order/sections/slide size
  - a flattened `shapes` table (slide, id, type, text, geometry, font, color, placeholder) for DuckDB

This lets the agent reason over a whole deck **without** dumping it into the model: `rg` finds every
shape citing a stale figure; `jq`/DuckDB lint the deck in one pass (off-brand fonts/colors, text
overflow, missing alt text, slides over a bullet budget, inconsistent terminology); `diff`/`sha256`
catch duplicate/near-duplicate slides and version drift. The model gets exact shape ids back and emits
a **batch** of surgical, reversible edits under one approval — formatting-preserving, not regenerate.

## Command grammar (typed; dovetails with the parallel grammar expansion)

Add **Layer-B read commands** to `command-grammar.ts` (with the existing `outline/read/search/context`)
— all typed and Zod-validated, no free-form shell:

- `query "<sql>"` — the hardened DuckDB tool over seeded tables.
- `probe <file> <tool> [args]` — runs ONE allow-listed safe tool
  (`head/tail/wc/rg/jq/csv-profile/diff/sha256`/outline) over a workspace file. NB: `inspect` is
  **already taken** by the expanded host-read grammar (`inspect <selector>` reads a host context ref),
  so the workspace safe-tool verb is `probe`, not `inspect`.
- Context seeding goes through the existing **`context` verb / context family** in the expanded grammar
  (alongside `list/properties/comments/attachments/tables/slides/neighbors/open`); the workspace is one
  consumer. The model may *request* seeding; the **runtime decides whether it is allowed**.

A new `analyze` intent (output `chat`) advertises a `compute` capability via capability-closure
(ADR-0006) **only when** the workspace is available and the host exposes a seedable representation —
tabular data (Excel) or a structured deck (PowerPoint). Skills choose
from the **runtime capability manifest** (the parallel "skills consume the manifest" change), so the
planner never offers compute where the host can't support it.

The Python parsers (`skill/scripts/parse_commands.py`) stay in lockstep with the TS grammar
(CLAUDE.md; TS is authoritative).

## Surgical ops (v1, through the gate)

Complete/add bridge actuate handlers for kinds the contract already models. Each is an instance of the
parallel **`create`/`update` families** and MUST satisfy the conformance matrix (below), with capability
advertisement, an `InverseDescriptor` (reversibility), and the provenance stamp (content-hash + sources
+ identity + timestamp). The gate path is unchanged.

**Excel (`bridge-excel`):**
- `insert-pivot` → **native Excel PivotTable** (`worksheet.pivotTables.add`, ExcelApi 1.8+); fallback
  to a computed values table when unsupported (degraded, not silent).
- `apply-conditional-format` (`cf` verb — cellValue/dataBar/colorScale/top, already in
  `ActuationParamsSchema`).
- `insert-table` / `format-range` (number formats, bold, borders); computed columns reuse `set`/`spill`.

**PowerPoint (`bridge-powerpoint`):**
- `set-shape-text` (already landed: target.slideId/shapeId + text, `restore-text` inverse).
- `format-shape` / `set-shape-geometry` — recolor, font/size, reposition/resize a shape by id (the
  surgical, formatting-preserving edits the deck-audit pipelines feed into).
- `insert-slide` (already present) for net-new slides.
- Edits are addressed by the same stable slide/shape ids `exportDeck()` emits, so a tool finding maps
  1:1 to a reversible write. (Native pivot is Excel-only; PowerPoint's "100x" is search/audit + batched
  surgical shape edits, not a pivot engine.)

## Trust & guardrails

- Worker: no-network, no Office.js, no DOM. DuckDB locked as above. Per-query timeout + memory cap.
- VFS: memory-only, byte caps, MIME allow-list, content hashes, lifecycle; no durable content/token
  persistence.
- Host data entering the workspace is **untrusted** → still Model-Armor-screened as it enters model
  context; tool/SQL output re-enters as data, never as instructions.
- `security-reviewer` MUST run before this is marked done (guardrails + a new execution surface).

## Testing & conformance

- `compute`: unit tests for `Sandbox` (`query`/`probe` over fixtures), VFS caps/hash/lifecycle, and
  **isolation proofs** — no general shell/process/eval; DuckDB cannot `INSTALL`/`LOAD`, reach the
  network, or read/write host files; queries are read-only; nothing persists.
- `bridge-excel`: actuate tests for `insert-pivot`/`cf`/`insert-table`/`format-range` incl. inverse
  (reversibility) via the fake-Excel harness; `exportTable` snapshot shape.
- `bridge-powerpoint`: `exportDeck` representation (slides/shapes/theme keyed by stable ids) +
  actuate tests for `set-shape-text`/`format-shape`/`set-shape-geometry` incl. inverse, via the
  fake-PowerPoint harness. A deck-scale audit fixture (`rg`/`jq` over an exported deck → batched edits).
- **Conformance matrix** (shared with the parallel work): every advertised command has a row asserting
  it wires parser → compiler → (for effects) ActuationRequest → bridge handler → dry-run → preview →
  approval → policy gate → actuation → provenance. Read verbs (`query`/`probe`) assert the read row
  (parser/compiler/handler) only.
- `web-shell`: a fixture-driven analyze → transcript → gated-write integration test.

## Consequences

- Real tabular analysis in the browser without weakening the trust model; pivots/surgical edits become
  reversible, provenanced host writes.
- Adds a WASM payload (DuckDB ~3–6MB) — lazy-loaded; feature-flagged off on legacy IE/EdgeHTML webviews
  and constrained mobile, where compute degrades to "unavailable" cleanly.
- A new code-bearing surface exists; mitigated by the lockdown + isolation tests + security review.

## Out of scope (later phases)

- Word surgical-op breadth (separate phase; same matrix).
- A heavy Python/Pyodide tier (only if real ML/scientific-Python demand appears; explicit governance
  approval required, per the clean-room doc).
- Cross-surface workspace persistence.
