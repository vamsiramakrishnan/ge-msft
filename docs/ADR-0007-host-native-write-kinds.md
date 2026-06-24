# ADR-0007 — Host-native write kinds: generalizing the actuation model beyond the tracked-change template

**Status:** Proposed (2026-06-24) · refines ADR-0004 (command protocol), ADR-0005 (composable algebra), ADR-0006 (capability closure). Scope: the in-document write surface (Plane A). Estate writes (Plane B) and cross-surface composition are out of scope here.

## Context

The add-in's write contract today is a **narrow, scalar/text mutation set**: `write-cells`,
`format-cells`, `add-comment`, `comment-reply` (Excel), `tracked-change` (Word), `insert-slide`
(PPT), `append-page` (OneNote), `reply-mail`/`create-mail` (Outlook), `post-message` (Teams). Every
admitted write satisfies three invariants (CLAUDE.md; `docs/02-design.md:51`): it is **anchored** to a
re-resolvable host location, **provenanced** with a content hash in durable host metadata, and
**reversible + legible** as one gated line.

Those invariants are load-bearing — they are the trust model for an untrusted-model, client-direct
architecture (ADR-0001). But the *mechanism* we chose for them was templated too literally on Word:
**"a write is a tracked-change-shaped scalar/text edit."** Anything that isn't scalar/text fell off
the bed. The result (`docs/CAPABILITY-MAP.md:125-145`) is a large band of host capability the surfaces
*expose* and the invariants would *permit*, sitting unbuilt:

- **Excel:** charts, tables, pivot tables, conditional formatting, named formulas, entity cards
  (`set-entity-card` is modeled but unadvertised), most font facets.
- **Word:** `insert-text`, `replace-selection`, `insert-ooxml`, `fill-content-control` — all modeled,
  none advertised.
- **PowerPoint:** shapes, `set-speaker-notes` (modeled, unadvertised).

The diagnostic insight: a chart **is more reversible than the text rewrite we already allow** — it is
a *derivative view* of a range the user already approved, deletable by name, and surfaces no new
information. It was excluded for *shape*, not *risk*.

## Decision

**Generalize the actuation from a fixed tracked-change template into a per-kind write *strategy*.**
An actuation is admitted iff it can supply four things; the *mechanism* for each becomes a per-kind
strategy, not one universal template:

> A `WriteStrategy` declares, per `ActuationKind`:
> 1. **Anchor** — how the effect is targeted + re-resolved at apply-time (a range, an object name, a
>    slide index).
> 2. **Inverse** — a recorded descriptor that undoes it (`delete table by name`, `restore prior
>    format`, `clear CF rule`). *This is the new first-class concept* — reversibility becomes an
>    explicit recorded operation, not an implicit property of tracked-changes.
> 3. **Provenance** — the agent/identity/sources/hash record, keyed by the anchor, written to the
>    host's durable store (Excel workbook settings; `provenance-record.ts`).
> 4. **Preview** — the dry-run effect the approval card renders (ADR-0005 plan-level gate).

The fail-closed gate, the closure conformance test (ADR-0006), formula safety, and untrusted-content-
as-data are **unchanged** — every new kind flows through the same machinery. We are widening *what*
can be a gated effect, not loosening *how* effects gate.

### What this is NOT

- **Not new planner intents.** The seven Copilot-altitude verbs (`ask · summarize · explain · rewrite ·
  review · draft · notes`) are unchanged. "Chart the Q3 numbers" is a `draft`/`rewrite`-altitude
  intent that *compiles down* to the new executor verbs. The planner/`parse_plan.py` side does not
  change — only the executor grammar and the bridges.
- **Not silent or auto-applied.** Every new kind is gated and reversible; nothing auto-sends.

## The plan, by the five questions

### 1 · Capability map — the new kinds

New `ActuationKind`s in `packages/contracts/src/capability.ts`, with their `WriteStrategy` (anchor /
inverse), grouped by how cleanly they fit:

| Kind | Surface | Office.js | Anchor | Inverse |
|---|---|---|---|---|
| **`create-table`** | Excel | `worksheet.tables.add(range, hasHeaders)` | range + table name | `table.delete()` |
| **`insert-chart`** | Excel | `worksheet.charts.add(type, sourceRange, seriesBy)` | chart name + source range | `chart.delete()` |
| **`format-conditional`** | Excel | `range.conditionalFormats.add(rule)` | range + rule ordinal | clear the added rule |
| **`create-pivot`** | Excel | `workbook.pivotTables.add(name, source, dest)` | pivot name + dest | `pivotTable.delete()` |
| `spill-cells` | Excel | `range.values = grid` | range | restore prior values (snapshot in inverse) |

`spill-cells` is the composition keystone (see §3) — it is `write-cells` widened to accept a whole
**table Value** rather than a single scalar. **Tier 2** (cheap follow-ons, already modeled — just need
a bridge `actuate()` case + advertise per ADR-0006): Word `insert-text`/`replace-selection`/
`insert-ooxml`/`fill-content-control`, PPT `set-speaker-notes`, Excel `set-entity-card`.

New `ActuationParams` fields (all optional, additive — the existing schema only grows):

```ts
// capability.ts — ActuationParamsSchema gains:
table:  z.object({ range: z.string(), hasHeaders: z.boolean().default(true), name: z.string().optional() }).optional(),
chart:  z.object({ chartType: z.enum(['column','bar','line','pie','scatter','area']),
                   sourceRange: z.string(), seriesBy: z.enum(['rows','columns','auto']).default('auto'),
                   title: z.string().optional() }).optional(),
conditional: z.object({ range: z.string(),
                        rule: z.discriminatedUnion('kind', [ /* cellValue | colorScale | dataBar | top */ ]) }).optional(),
pivot:  z.object({ name: z.string(), sourceRange: z.string(), destRange: z.string(),
                   rows: z.array(z.string()), columns: z.array(z.string()).default([]),
                   values: z.array(z.object({ field: z.string(), agg: z.enum(['sum','count','avg','min','max']) })) }).optional(),
inverse: InverseDescriptorSchema.optional(),   // the recorded undo — persisted in provenance
```

### 2 · How the CLI verbs evolve

New executor verbs in `WRITE_VERB_TO_KIND` (`command-grammar.ts`) + their `ParsedCommand` variants.
They follow the existing arg grammar exactly (`format`'s `verb <anchor> key=value…` shape):

```
table <range> [headers]                         → create-table
  table Sales!A1:D200 headers
chart <type> <range> [title="…"] [series=rows|cols]   → insert-chart
  chart column Report!A1:B11 title="Top regions by revenue"
cf <range> <rule>                               → format-conditional
  cf Sales!E2:E200 >100000 fill=#C6EFCE
  cf Sales!E2:E200 databar
spill <range> = <expr>                          → spill-cells   (table Value → grid; see §3)
  spill Report!A1 = ($top)
pivot <name> <source> <dest> rows=… values=…    → create-pivot   (Phase E)
```

Design rules, consistent with ADR-0004/0005:
- **Anchor slots stay literal** (a range/name is never an expression) — same as `format`/`suggest`.
- **Text slots are expression-bearing** — `chart … title=( $anz | count )` parses a `titleExpr`,
  exactly like `comment`'s `textExpr`.
- **The data source is a range** (`chart`/`table`/`pivot` read a range), and a range is produced
  either directly or by `spill` (§3) — so composition feeds visualization without a magic
  table-as-source coercion.

`ParsedCommand` additions:
```ts
| { verb: 'table'; range: string; headers: boolean }
| { verb: 'chart'; chartType: string; range: string; title?: string; titleExpr?: ParsedExpr; seriesBy?: 'rows'|'columns' }
| { verb: 'cf'; range: string; props: Record<string,string> }
| { verb: 'spill'; range: string; valueExpr: ParsedExpr }
| { verb: 'pivot'; name: string; source: string; dest: string; rows: string[]; values: string[] }
```

### 3 · How composition works (the interesting part)

ADR-0005's value model is `Table | Number | Text`. Today a `Table` can only terminate as **slide
bullets** or be **rejected** for a scalar slot (`set` refuses a table). The unlock is a **new table
sink: table → cell grid.** That single coercion turns "analyze → shape → materialize → visualize"
into one pure pipeline that dead-ends in gated effects:

```
let $anz = read Sales!A1:D5000 | filter Quarter=Q3 | select Region,Revenue | sort Revenue desc | head 10
spill  Report!A1 = ($anz)          # NEW: table Value → write-cells grid (params.cells = rows)
table  Report!A1:B11 headers       # promote the spilled grid to a native Table
chart  column Report!A1:B11 title=("Top regions — " ~ ($anz | count) ~ " of N")
done
```

Mechanics, reusing the existing dry-run engine (`assist-session.ts` `resolveEffectArgs`/`evalExpr`):
- `spill`'s `valueExpr` evaluates to a `Table` Value at dry-run; **a new `valueToGrid()` coercion**
  (mirror of the existing `valueToBullets`) renders it to `params.cells: string[][]`. The existing
  formula-safety screen (`isUnsafeFormula`) runs on every spilled cell — composed data can never
  promote to executable active content.
- **`spill` returns its target range as a bindable address**, so the dynamic size is composable:
  `let $r = spill Report!A1 = ($anz)` then `table ($r) headers` / `chart column ($r)`. The
  dry-run knows the table's dimensions, so the resulting range (`Report!A1:B11`) is computed before
  the dependent `table`/`chart` resolves. This keeps the chain pure and gated — every intermediate
  range is visible on the approval card, nothing is hidden scratch state.
- Charts/tables therefore consume a **range produced by composition**, not a table-as-source magic
  coercion — which keeps anchor + inverse + provenance well-defined (you anchor/undo on the range
  and the named object, both concrete).

Net: composition gains exactly one primitive (`table → grid`), and that primitive is what makes the
whole visualization band composable. No change to the transform set, the binding env, or the gate.

### 4 · How the skill is updated (parity)

The TS side is authoritative; the skill mirrors it (`parity_corpus_test.py` enforces the verb-set
equality, `parity_test.py` enforces the planner intents — *unchanged here*). For each new executor
verb:

- **`m365-surface-commander/scripts/parse_commands.py`** — add to `WRITE_VERBS` and a dispatch arm in
  `parse_line()` (the `format`-style positional+kv parse for `table`/`chart`/`cf`/`pivot`; the
  `set`-style anchor+expr parse for `spill`).
- **`references/command-grammar.md`** — add the verb rows to the write table, and extend the
  Composition section with the **table → grid** sink (today it documents table → bullets / scalar).
- **`references/capability-map.md`** — add the new verbs to the per-app writes table (Excel column).
  `parity_corpus_test.py` extracts this table and asserts `ALL_VERBS == documented` — it fails the
  build until the doc and the parser agree.
- **`golden-corpus.jsonl`** — add cases per verb: a valid line, an error line (bad args), and one
  **composition case** (`spill` of a filtered table → grid) plus the fail-closed `done`-with-error
  invariant for a malformed new verb.
- **`SKILL.md`** — add the verbs to the quick reference; re-bundle + re-upload via the create tooling.
- **Planner side is untouched** — these are executor verbs the seven intents compile down to.

### 5 · Bringing it together — phasing

One ADR (this), then four build phases driven from `BUILD-PLAN.md`, each independently shippable and
each ending green (typecheck · test · lint · closure conformance · security-reviewer where it touches
provenance/guardrails):

- **Phase A — contract + composition core (pure, no host).** New `ActuationKind`s + params; the
  `WriteStrategy`/`InverseDescriptor` types; the `valueToGrid` coercion + `spill` verb in the grammar
  and `compileCommand`; `resolveEffectArgs` handling for `titleExpr`/`spill`. Unit-tested in
  `contracts` + `runtime`. *Nothing actuates yet.*
- **Phase B — Excel bridge.** Implement `create-table`, `insert-chart`, `format-conditional`,
  `spill-cells` in `excel-bridge.ts`/`actuate-plan.ts`: planner (pure, tested) + applier (Office.js),
  each recording its **inverse** into the provenance payload, advertised via `HANDLED_ACTUATIONS`.
  Closure test goes green. **security-reviewer** (untrusted data → chart titles/labels; formula safety
  on spilled cells; the inverse descriptor can't be coerced into deleting an un-owned object).
- **Phase C — skill parity.** Mirror the verbs in Python, grammar docs, golden corpus; update the
  parity tests' expected verb set. `parity_corpus_test.py` green.
- **Phase D — UX.** Quick actions ("Chart this range", "Make it a table", "Highlight outliers") built
  on the H typed-parameter forms (chart type, target range); approval-card previews for the new kinds
  (the dry-run shows the source range + the object to be created + the recorded inverse).
- **Phase E — pivots + Tier 2.** `create-pivot`; advertise the already-modeled Word
  (`insert-text`/`replace-selection`/`insert-ooxml`/`fill-content-control`) and PPT
  (`set-speaker-notes`) kinds now that the per-kind strategy pattern exists.

### Invariants preserved (the part we do not touch)

Fail-closed plan/approval gate · closure conformance per surface · provenance-or-observable-drop
(`provenanceDropped`/`provenanceMissing`) · formula safety on every written/spilled cell · untrusted
host content (and composed data) framed strictly as data, never instructions · no auto-send.

## Consequences

- **Unlocks the Procrustean band** (charts/tables/CF/pivots + the modeled-unadvertised Word/PPT kinds)
  without weakening the trust model — reversibility becomes an explicit recorded inverse rather than
  an implicit tracked-change property.
- **One new composition primitive** (`table → grid` via `spill`) makes the whole analyze→visualize
  flow a pure, gated pipeline.
- **Cost is real and per-kind:** each kind owns an anchor + inverse + applier + tests; this is "deeper
  host operations," not "one more verb." Pivots are the most involved (self-restructuring; the
  inverse is delete-by-name, but re-resolution semantics need care).
- **Estate writes and cross-surface composition remain out of scope** — they need separate
  authorization (Plane B OAuth scopes) and a multi-bridge router, tracked elsewhere.

## Security notes carried forward (from the Phase A review)

Phase A is pure (no host) and weakens no invariant — the `spill` grid provably reuses the unchanged
`write-cells` formula screen (`isUnsafeFormula`, whole-write degrade), `SPILL_ROW_CAP` fails loud
(corrective, never a silent truncation), unadvertised kinds are type-check-rejected before the gate,
and every new verb stays a gated pipeline terminal. Two forward-looking items the later phases MUST
honor:

- **Phase B — inverse identity.** A `delete-object` `InverseDescriptor` carries a bare `name`. The
  undo applier MUST verify the object was the one *this change minted* (match the name recorded at
  apply-time, scoped to the provenance entry) and degrade on a mismatch — it must NEVER re-resolve an
  arbitrary name against the live workbook, or an undo could delete a hand-made table/chart. The Phase
  B security-reviewer pass must assert this explicitly. Likewise, the bridge must keep every new Excel
  write (incl. spilled grids) routed through `splitFormulaGrid`, never a separate applier that skips
  the screen.
- **Phase C — quote parity.** `tokenizeArgs` does not honor `\"` escapes the way `scanQuoted` does (a
  cosmetic parse-differential on a chart title). Align it with `scanQuoted` when the skill-side parser
  is mirrored.

## Open questions (need a decision before Phase A)

1. **`spill` vs. overloading `set`** — a distinct `spill` verb (table→grid) keeps `set` scalar-only
   and legible; overloading `set` is fewer verbs but blurs the scalar/grid line. *Recommend a distinct
   `spill`.*
2. **Chart source — range-only vs. direct table-expr** — range-only (materialize via `spill` first)
   keeps anchor/provenance concrete; direct `chart (…expr…)` is terser but needs hidden scratch state.
   *Recommend range-only for Phase A–D; revisit direct-source later.*
3. **Inverse persistence depth** — store just "delete object N" (cheap, reverses creation) vs. a full
   prior-state snapshot for in-place kinds like `format-conditional` (heavier, true undo). *Recommend
   snapshot only where the op mutates existing state (CF, spill-over-data); delete-descriptor for pure
   additions (table/chart/pivot).*

## Cross-surface expansion (the locked catalog)

The Excel host-native kinds (`create-table`/`insert-chart`/`format-conditional`) were the first
instance of a general pattern: **every surface exposes powerful host-native writes we had not
modeled.** That full map is now locked in **`docs/CAPABILITY-CATALOG.md`** (the authoritative,
typing-grounded catalog) and in the contract (`ActuationKindSchema` now spans ~85 kinds across all six
surfaces). The catalog was produced by inventorying each surface against `@types/office-js` + Microsoft
Graph; no kind is listed without a concrete API behind it.

The expansion does **not** change this ADR's thesis — it generalizes it. Two planes:

- **Plane A (client-direct)** — host writes via Office.js/TeamsJS in the open document/draft. These fit
  the existing model directly: anchored + reversible (a recorded `InverseDescriptor`) + provenanced +
  gated. The `InverseDescriptorSchema` grew the ops these need (`restore-text`/`restore-style`/
  `restore-shape-format`/`restore-slide`/`move-slide`/`apply-layout`/`detach-list`/
  `restore-content-control`/`restore-doc-properties`/`restore-mail-state`/`not-reversible`), and
  `delete-object.objectType` widened to the full object set. The **inverse-identity security rule above
  applies to all of them**: an undo deletes/restores only the object *this change minted*, never an
  arbitrary host object re-resolved by name.
- **Plane B (estate)** — writes to the M365 estate via Microsoft Graph (mailbox triage, calendar,
  Teams channel posts, server-side notebook edits). These need delegated `*.ReadWrite`/`*.Send` scopes
  **and** a `GraphClient` write path, neither of which exists today (the client is read-only by
  design). They are catalogued but **deferred to a dedicated estate-write ADR** — they lack the
  on-send safety net, so they need a distinct per-call confirm gate, and the highest-privilege ones
  (`create-mail-rule`, `send-activity-notification`) need elevated, explicit warnings.

**Explicitly out of scope / not modeled** (documented in the catalog, not in the enum): Graph
`sendMail` (escapes the on-send gate — violates no-auto-send), Teams *interactive* Adaptive Cards
(need Bot Framework server credentials — violates client-direct), and PPT `set-speaker-notes` (no host
write path in current typings). Honesty about these boundaries is part of locking the map.

Phasing for the expansion lives in the catalog's "Phasing" section: Plane A per surface
(Excel → Word → PowerPoint → OneNote → Outlook-compose), then Plane B behind the estate-write ADR.
