# Command reliability: most "random" failures were ours, not the model's

## Bottom line

**The add-in failed the same request in different ways because shared code gave Gemini too little
to work with, then accepted what it guessed.** In 8 live chart requests on 2026-09-30, none would
have produced the right chart. Four identical prompts produced four different `chart` commands.
Gemini does vary from run to run, and streamAssist exposes no temperature control. But five of the
six causes below are deterministic defects in shared packages (`content`, `contracts`, `runtime`,
`web-shell`, `gemini-client`). Every surface runs through those packages, so fixing them once
fixes Excel, Word, PowerPoint and Outlook together.

This change fixes A, B, D, E and G, Excel charts under C, and the sign-in, PowerPoint start-up and
Word write failures the cross-surface run found. Typed targets for the other write kinds (the rest
of C), F and H are open.

## Evidence

**The client test sheets show use cases that failed and then passed with no code change.** Excel
"Format cells", Word "Apply style" and "Native table", and Outlook "New email draft" and "Set
subject / body" are each marked "NOW WORKING - NOT WORKING BEFORE … No code change was needed"
(`client-{excel,word,outlook}-tab.csv`, validation of 2026-09-23 and 2026-09-24).

**A live run reproduced it: 0 of 8 chart requests would have produced a correct chart.** This was
Excel for the web, with an orders table in `Sheet2!A1:J11` (a header row and 10 data rows). No
skills were mounted and the engine's default model was used. Every staged write was rejected, so
nothing was applied.

| Run | Prompt | Route | Model's final command | Outcome |
| --- | --- | --- | --- | --- |
| r1 | Create a bar chart of Total by Product | planner → executor | `chart bar {"source":"Sheet2!A1:J11",…}` | Approval card; would fail inside Excel after approval |
| r2 | same | planner → executor | `analyze {"action":"chart",…}` + `done` | Ended incomplete; no retry |
| r3 | same | planner → executor | `chart bar Sheet2!C1:C10,Sheet2!G1:G10` | Approval card; drops row 11; invalid two-area address |
| r4 | same | planner → executor | `chart {"type":"bar","axes":…}` + `done` | Ended incomplete; no retry |
| v1 | show me a bar chart of total by product | chat | none | Claimed "I have created the bar chart"; named a product not in the sheet |
| v2 | I want a bar chart of totals per product | chat | none | Invented totals (Laptop $4,800; the real total is 130,000) |
| v3 | bar chart of Total by Product | chat | none | Said no data was provided |
| v4 | Create a column chart of Total by Region titled Revenue by region | planner → executor (4 turns) | `chart Sheet2!L1 type=column data=$summary.range …` | Ended incomplete |

The executor programs from r1–r4 and v4 are stored verbatim in
`packages/runtime/src/replay/live-2026-09-30-excel-chart.json` and replayed by the tests.

## Root causes

**A. The executor was never shown the syntax of the write it had to emit.** The bootstrap prompt
listed a hand-written set of signatures. The other writes, `chart` among them, were listed by name
only. It then added the two "relevant command" cards that ranked highest on the task text. After
the planner, that task text starts with a fixed wrapper: "read live host content…", "the open
Microsoft 365 surface". Those words gave `read` and `open` full name matches, and they beat `chart`
on grammar order. The model saw `chart` with no signature and invented one. The specialised
`/<kind>` commands were worse off: their usage is only `/<kind> [key=value ...]`, which names no
parameter. All surfaces. (`runtime/src/command-protocol.ts` bootstrap,
`runtime/src/capability-catalog.ts` `discoverCommands`, `web-shell/src/controller.ts`
`renderConfirmedPlanTask`.)

**B. The document snapshot hid the addresses and columns the model needed.** `renderDocState`
never printed an inventory entry's `id`, so the model saw no A1 address. It clipped a table's label
to 64 characters (`"| Order ID | … | Unit Pri"`), so the `Total` column was invisible. It
summarised the table as `10 rows × 10 cols`, and the model read that as rows 1–10, so row 11 was
lost. The planner added "calculate Total column if not present" in every planned run because it
could not see that column. All six bridges build their snapshot through this one function.
(`content/src/doc-state-builder.ts`.)

**C. Invalid targets pass all checks until after the user approves.** `insert-chart`'s
`sourceRange` is typed only as `z.string()` (`contracts/src/capability.ts`), so the JSON blob from r1
reached the approval card. It fails only inside `Excel.run`. `parseAddress` splits at the last `!`,
so r1 names a sheet `{"source":"Sheet2`, and r3's two-area address names a sheet
`Sheet2!C1:C10,Sheet2`. `DocBridge` has no read-only preflight, so no bridge can check a target
before approval. All surfaces. **Partly fixed in this change, for Excel charts; typed targets for
the other write kinds and a bridge preflight are open.**

**D. One invalid line followed by `done` ended the task before the model saw its error.** The loop
treated `done` as final even when earlier lines in the same program had failed. It then marked the
task incomplete and never sent the errors back. Runs r2 and r4 ended this way. Models append `done`
almost every time. All surfaces. (`runtime/src/assist-session.ts` `executeProgramTurn`.)

**E. The wording, not the capability, chooses the route, and the chat route cannot write.** Per-surface
regexes in `web-shell/src/taskpane/components/App.tsx` only promote requests that start with an
imperative verb. "show me…", "I want…" and "bar chart of…" went to chat, which sent only the
empty selected cell, so the model invented data and in one case claimed success. **Fixed in this change: routing, and chat turns now read the document.** The cross-surface
run showed the same for the client sheet's own prompts: "In G12 put a formula that totals G2:G11"
went to chat 3 times out of 3, and chat answered with instructions for the user to type
`=SUM(G2:G11)`. The router's tests show the verb list being extended one reported phrase at a time.

**G. The planner asks a question whenever a target is not in the snapshot.** On a 66-paragraph Word
document, "Apply the Heading 2 style to the line 'Payment terms'" produced a clarification: the
line "was not found in the current document outline". The snapshot lists structure (headings,
tables, slides), not every line, and the planner is not allowed to read. So any plain-text target
in a long document triggered a question.

**F. The environment adds variance of its own.** Skills are disabled in this deployment, so the
add-in's own prompts (the command bootstrap and the planner prompt in `contracts`) are the only
instructions the model gets. That is why fixes A and B matter so much. `VITE_GE_MODEL_ID` is
blank, so the engine's default model is used, and Google can change it. **Not fixed in this
change.**

## Fixes in this change

### B — the snapshot names the addressable ref, every column, and the header row

- **Each inventory entry now carries a `ref` a command can use.** Excel ranges show the bare A1
  address (`ref="Sheet2!A1:J11"`). The bridge rejects the `range:` locator prefix, so it is
  stripped. Other host locators (`slide:<id>`, `cc:<id>`) are shown verbatim. Ids the builder
  invents (`heading:N`, `table:N`) get no `ref`, because no command accepts them.
- **A table's label is its full header row** (up to 480 characters) instead of a 64-character clip.
  A wider header ends in `…`, so the model can tell that columns were cut.
- **The summary names the header row**: `header row + 10 data rows, 10 cols`.
- **The untrusted-data boundary is unchanged.** `ref` is host text (sheet names are typed by
  users), so it goes through `safe()` and is quoted like every other host value. A test checks that
  a sheet name containing `</doc_state>` cannot close the envelope.

Code: `content/src/doc-state-builder.ts` (`tableSummary`, `tableTitle`, `commandRef`, inventory
rendering). Tests: `content/src/doc-state-builder.test.ts`.

### A — every write is disclosed with its exact syntax; cards rank on the user's request

- **The bootstrap lists every advertised write, derived from the capability manifest.** No
  hand-written list is involved. Core verbs show their usage line. Specialised `/<kind>` commands
  show their registry example, for example `/format-shape slide=1 shape=title fill=#0F6CBD
  fontColor=#FFFFFF`. Only a one-line example of the same command is used; otherwise the usage
  line is shown. The largest surface prompt stays under the existing 4 KiB budget, and a test
  asserts that for every surface.
- **The prompt says that snapshot values are entity-escaped.** A sheet named `P&L` renders as
  `P&amp;L` inside `<doc_state>`. The escaping stays, since it protects the envelope, but the model
  is told to write the plain characters in commands.
- **Relevant-command cards rank on what the user asked.** `commandIntentText(task)` returns only a
  confirmed plan's request, steps and exclusions, and returns any other task unchanged. The
  `<confirmed_plan>` delimiters are now `contracts` constants shared by the wrapper
  (`web-shell/src/controller.ts`), its display parser (`MessageThread.tsx`) and the ranking. Task
  text still only influences ranking; it is never copied into instructions.

Code: `contracts/src/command-plan.ts` (`commandIntentText`, `CONFIRMED_PLAN_OPEN/CLOSE`),
`runtime/src/capability-catalog.ts` (`writeSignatures`), `runtime/src/command-protocol.ts`
(`renderCommandBootstrap`). Tests: `runtime/src/replay/command-reliability.replay.test.ts`,
`contracts/src/command-plan.test.ts`.

### D — a program with a failed line applies nothing and goes back for a bounded repair

- **If any line in a model-authored program fails to parse, compile, resolve or read, none of its
  writes run and `done` is ignored.** The errors, plus an instruction to re-emit the complete
  corrected program, go back to the model on the next turn. The pane shows a `repair` step.
- **All-or-nothing prevents double writes.** The repair turn re-emits the whole program. If a valid
  write from the failed program had already applied, it would apply twice, which is the "applied
  edits twice" failure in the Word test sheet.
- **Policy notes do not trigger a repair.** Cap and budget refusals and "done cannot be batched
  with a write" are marked as advisories, so a valid write still proceeds under the existing rules.
- **A repaired program leaves no bindings behind.** Its `let` and `analyze` bindings are rolled
  back, so the corrected program can reuse the same names. Live run a-r2 failed on exactly this:
  the corrected program's `let $data` collided with the failed program's `$data`.
- **An effect that already landed is named, not hidden.** An approved `share` takes effect inline
  and cannot be withheld. The repair instruction and the pane list it as already applied, so the
  model does not repeat it. A share's outcome, including a user's denial, never triggers a repair.
- **The repair is bounded by consecutive failing turns.** Three turns in a row with a failed line
  (the first attempt plus 2 corrections) stop the task as `incomplete` with `repair_exhausted`,
  whether or not those turns staged a write. Counting only turns with a staged write let live runs
  a-r2 and a-r4 repeat one mistake for all 12 turns.
- **A corrective error says how to fix the mistake.** "Unknown artifact binding $x" now says that
  analyze inputs come from `let $x = analyze {"kind":"capture",...}` and that a `read` result cannot
  be one. In a-r2 the model repeated that mistake 10 times.
- **Out of scope by design:** `finish when=verified` keeps its documented stop-without-replay rule
  (`docs/COMMAND-PERFORMANCE.md`). Direct programs (`mode: 'program'`) and SDK analysis programs
  have no model to repair them and behave as before.

Code: `runtime/src/assist-session.ts` (`executeProgramTurn`, `runCommandsCore`,
`isValidationFailure`, `advisory`, `repairInstruction`, `MAX_REPAIR_TURNS`),
`runtime/src/analysis-program.ts` (`snapshot`, `restore`, the corrective message),
`web-shell/src/controller.ts` (`repair` step). Tests: `runtime/src/replay/command-repair.replay.test.ts`. Two integration tests
in `web-shell/src/taskpane/grammar-effect-parity.integration.test.ts` pinned the old partial
application; they now script the model's corrected turn and still assert that the rejected line
never reaches the bridge.

### E — every non-question goes to the planner

- **The router now defaults to the planner on any surface that can write.** Only text that reads as
  a question goes straight to chat: it starts with a question word ("which", "what", "how",
  "summarize", "explain", "can you tell me…") or ends with `?`. The existing action pattern still
  wins, so "Could you create a chart?" goes to the planner.
- **The planner can still answer in chat.** If it classifies the request as a question, the task is
  sent to chat as before, so a misrouted question costs one extra model call, not a wrong result.

Code: `web-shell/src/taskpane/components/App.tsx` (`shouldUsePlannerForFreeText`). Test:
`web-shell/src/taskpane/command-surface.integration.test.ts` pins the route of every client-sheet
prompt and the three live chart wordings.

### C, charts — refuse malformed chart sources before approval, and chart separate areas

- **A chart's `sourceRange` must now be A1 areas on one sheet or a defined name.** The contract
  (`CHART_SOURCE_PATTERN`) refuses anything else during the dry run, so fix D sends it back for
  repair before the user sees an approval card. All four malformed forms the model produced live
  (a JSON object, `Sheet2!L1 type=column`, a dangling `Sheet2`, `$summary.range`) are refused.
- **The Excel bridge charts non-adjacent columns.** For "a bar chart of Total by Product" the model
  correctly asks for `Sheet2!C1:C11,Sheet2!G1:G11` (Product and Total), which `Worksheet.getRange`
  rejects. The bridge now charts the first value area (its header becomes the series name), sets the
  first area as the categories without its header, and adds a series per further area. This needs
  ExcelApi 1.7; an older host gets an explicit `unsupported_host` result. Areas on different sheets
  are refused as `invalid_target`.

Code: `contracts/src/capability.ts` (`CHART_SOURCE_PATTERN`), `bridge-excel/src/actuate-plan.ts`
(`chartAreas`), `bridge-excel/src/excel-bridge.ts` (`addAreaChart`). Tests:
`bridge-excel/src/excel-bridge.test.ts`, `runtime/src/replay/command-repair.replay.test.ts`
(replays r1's JSON source).

### Bridge reads that failed silently on the real hosts

The cross-surface run found two reads that worked against the test fakes and failed on Office for
the web, where a swallowed error looked like "nothing found".

- **Word search never returned a match on a branch without `24a4e81`.** `searchText` read
  `results.items` before `ctx.sync()`. Word throws `PropertyNotLoaded` there, the surrounding
  `catch` returned `[]`, and the fake Word host handed items back immediately, so no test noticed.
  Chat questions about the document ("Which paragraph mentions the payment terms…") scored 0 of 3
  live, then 3 of 3 once the read synced first. Commit `24a4e81` on `merged-ge-fixes` had already
  made the same fix; this change keeps it, bounds the hits before loading a paragraph for each, and
  searches `^` literally. Both Word fakes now throw `PropertyNotLoaded` for `items` read before a
  sync, like Word does; with the old read, 5 tests fail.
- **A whole-document `read` returned headings only.** On Word it rendered the structure snapshot,
  so a model asked to comment on or rewrite body text searched for it for all 12 turns. A
  selector-less `read` now returns the bridge's whole-item context (Word document, PowerPoint deck,
  Outlook mail item, OneNote page) when the bridge lists one.
- **`done` in the same block as a read ended a task with nothing done.** A Word rewrite emitted
  `read` then `done`; the loop accepted it before the model saw what it read. Like `done` after a
  write, it is now refused and the read result is returned first.
- **Reading a numbers-only Excel range threw.** `splitHeaderRows` called `.trim()` on each cell,
  but Excel returns numbers, so `read Sheet2!G2:G11` failed with "r.trim is not a function". The
  model then commented the wrong cell (G4, not the lowest total in G8). Cells are now read as
  values. A range whose first row is numeric is no longer given that row as its header; its columns
  are named by sheet letter and each row carries its sheet row number, so the model can name the
  exact cell.
- **The bootstrap says pipelines cannot build targets.** In another run the model tried to compute
  the address (`comment Sheet2!G$target_row`). The prompt now says to choose the target from the
  read result and write it literally.

Code: `bridge-word/src/host-port.ts` (`searchText`), `bridge-excel/src/capture.ts`
(`splitHeaderRows`, `headerlessTable`), `runtime/src/assist-session.ts` (`wholeItemRead`, `done`
handling), `runtime/src/command-protocol.ts`. Tests:
`bridge-word/src/host-port.test.ts`, `web-shell/src/test-harness/fake-word.ts`,
`bridge-excel/src/capture.test.ts`.

### Sign-in stopped working until the cached id token expired

- **Every request failed with `WIF token exchange failed (400): invalid_grant … ID Token issued at …
  is stale to sign-in`.** Google's STS refuses an Entra id token issued too long before the
  exchange, even when it has not expired. MSAL's silent acquire kept serving that cached token, so
  every task failed until sign-in, and reloading the task pane did not help (live 2026-10-01).
- **On that error the exchange now asks MSAL for a newly issued token (`forceRefresh`) and retries
  once.** Live, the next exchange returned 400 and the retry 200. Other 400s are not retried. A failed
  exchange is a typed HTTP 4xx error, so outer retry policies do not repeat it. The client still
  holds only the user's own tokens, in memory; a security review of this change found nothing above
  low severity, and its low findings are fixed.

Code: `gemini-client/src/wif.ts` (`exchange`, `WifExchangeError`), `web-shell/src/auth-client.ts`
(`getIdToken({ forceRefresh })`). Tests: `gemini-client/src/wif.test.ts`,
`web-shell/src/auth-client.test.ts`.

### Host and navigation fixes found by the cross-surface run

- **The PowerPoint task pane stayed blank at start-up.** Boot awaited `Office.auth.getAuthContext()`
  for an optional login hint, and on PowerPoint for the web that call never settles (Word answers at
  once). Boot now waits at most 2 seconds and continues without the hint; the pane then mounted in
  about 6 seconds. Code: `web-shell/src/taskpane/office-login-hint.ts`.
- **Word OOXML inserts always failed.** The model writes a bare `<w:p>…</w:p>` fragment, and
  `insertOoxml` requires a flat OPC package, so every insert ended `outcome_unknown`. A fragment is
  now wrapped in the minimal package; a full package passes through. Code:
  `bridge-word/src/actuate-plan.ts` (`toOoxmlPackage`).
- **On a long Word document the model could find a heading but not the text under it.** A search
  hit now carries up to two following paragraphs (capped at 600 characters; the hit's own paragraph
  is kept whole). Code: `bridge-word/src/host-port.ts` (`followingText`).
- **A chart over "A1:J11" plotted every column.** Asked for "a column chart of Total by Product from
  A1:J11", the model charted the whole table. The `chart` help card now says to chart only the
  label and value columns, listing separate areas for non-adjacent columns, with that as its first
  example. The skill manifest was regenerated with `bun run emit:language` to keep the drift gate
  green.

### Regression check against the fixes already on `merged-ge-fixes`

A review of every behaviour introduced by `571685e`, `24a4e81`, `0a1aeb4`, `bb111c0`, `c6cba57`,
`32b2482` and `8402bb4` found one regression, now fixed: the duplicate-write refusal from
`24a4e81` ("already applied … do not repeat it") was counted as a failed line by fix D, so a program
that repeated a landed write lost its other writes and could end `repair_exhausted`. It is now an
advisory, and two replay tests fail without the fix. Everything else was preserved; the PowerPoint
and Outlook bridges are untouched. All 612 tests in the 14 test files those commits touched pass.

### E — a chat turn carries the document data the question needs

- **Every bridge's search matched the whole question as one substring, so chat read nothing.**
  "Which region had the highest total revenue in A1:J11?" reached Gemini with the snapshot and no
  cell values. One live run answered "South, 1825"; the data says West, 243,000.
- **A chat turn now probes the document the way a bridge can serve.** It reads the references the
  question names (`A1:J11`, `Sheet2!B4`, `slide 2`) through the bridge's own `readRange`, searches
  the whole question and then up to 3 distinctive words, and, when all of that finds nothing, reads
  the addressable tables listed in the snapshot. Selector syntax stays the bridge's: `readRange`
  returns nothing for a selector it cannot address. All reads stay framed as untrusted data and
  bounded by the existing `maxReads`.

Code: `runtime/src/chat-reads.ts` (`searchTerms`, `referenceCandidates`),
`runtime/src/assist-session.ts` (`chatReads`), `content` (`commandRef`, shared with the snapshot).
Test: `runtime/src/replay/chat-reads.replay.test.ts`.

### G — the planner treats a missing target as something to locate

- **The planner prompt now says the snapshot is partial.** A target the request names or quotes
  that is absent from the snapshot is not ambiguous; the planner plans a step to locate it, and the
  executor searches the document. It still asks when the request itself leaves the target, value or
  destination open.

- **The planner prompt defines each intent by its effect, and offers only the surface's intents.**
  It was offered all seven with no meaning attached, read `rewrite` as "reword text", and labelled
  "In G12 put a formula that totals G2:G11" as `ask`. Now `rewrite` reads "change existing content:
  text, values, formulas, formatting, styles, tables…", and the prompt says a request that changes
  the document never uses ask, summarize or explain.
- **A chat-labelled plan whose steps change the document is staged as an edit, not sent to chat.**
  On that mislabelled plan the controller silently re-sent the request to chat, which replied "The
  formula =SUM(G2:G11) has been placed in cell G12". Nothing had been written. The planner's steps
  are the stated work; when they describe a change, the plan is shown for confirmation under the
  surface's edit intent (`rewrite` on Word and Excel, `draft` on PowerPoint and Outlook).

Code: `contracts/src/command-plan.ts` (`renderPlanPrompt`, `planDescribesChange`,
`editIntentFor`), `runtime/src/assist-session.ts` (`planCore`), `web-shell/src/controller.ts`
(`proposePlan`). Tests: `contracts/src/command-plan.test.ts`, `web-shell/src/controller.test.ts`
(replays the recorded planner reply).

### Outlook

Outlook could not be run live. `web-shell/src/taskpane/outlook-reliability.integration.test.ts`
drives the real Outlook bridge, runtime and controller over the fake mailbox. It checks that the
first command turn shows the exact `mail`, `compose` and `/set-subject` syntax, that an invented
reply verb followed by `done` is repaired into exactly one reply form, and that an incomplete
`compose` is repaired into exactly one new draft. The two repair tests fail with fix D disabled.

## Verification

**Replay tests (no model).** The recorded prompt and model programs are replayed through the real
parser, compiler and loop. Each D test fails when the fix is disabled and passes with it. The A
test reproduces the recorded ranking failure on the raw task (`read`, `open`) and asserts `chart`
wins on the user's request.

**Definition of done.** On `fix/command-reliability-v2` (based on `merged-ge-fixes`):
`bun run typecheck` clean; `bun run test` 2,829 passed, 16 skipped, 0 failed; `bun run lint`
clean; Python parity checks pass (71 golden cases).

**Security review.** The `security-reviewer` agent found no critical or high issues. It confirmed
that a withheld write is never prepared, previewed or actuated, that the write cap is unchanged,
that document content can only cause writes to be withheld, and that the `<doc_state>` envelope
still escapes and quotes every host value. It raised 2 medium and 5 low findings, all fixed in this
change: inline `share` effects and surviving bindings under repair (medium), share denial and cap
counted as failures, the counter reset, entity-escaped refs, the missing elision marker, and
unfiltered signature examples (low). One low finding is recorded, not fixed: a hostile workbook
with 60 wide tables can still make the snapshot large, because length caps apply before escaping.

A second review covered the later fixes (chat reads, routing, intent relabelling, chart areas,
reads). No critical or high issues. Its three medium findings are fixed: a defined name such as
`FY24_Q3` was split into two cell areas, a failed multi-area chart could be left with no undo, and
search matches in a headerless grid got the wrong sheet rows. Fixed lows: a probe-length cap, per-
probe error isolation, a cap on fallback reads, bounded Word hits and literal `^`, `|` and line
breaks escaped in table cells, the planner's surface checked before use, multi-area chart
dependencies, and sheet names read after the last `!`. Recorded, not fixed: up to four used-range
loads per Excel chat turn, and a chat reply containing a `cmd` block still hands off to the
executor without the planner card (all writes stay gated).

**Live runs.** The same 8 prompts before and after the fixes; see the next section.

## Live verification after the fix

**The user's own example now produces the correct chart every time.** "Create a bar chart of Total
by Product" failed in 4 of 4 identical runs before the fixes; on the final build it charted the
Total series by product in 3 of 3 runs, through the two-area source the model chooses for
non-adjacent columns (`Sheet2!C1:C11,Sheet2!G1:G11`).

**Cross-surface run on the final build.** The client test-sheet prompts, each run 3 times unless
noted, every approval card approved, and every result checked in the document itself (cells,
charts, paragraphs, comments, slides), not in the chat text.

**PowerPoint: 21 of 21.**

| Case | Result |
| --- | --- |
| What is on slide 2? | 3/3 |
| Draft one slide titled Q4 outlook with 3 bullets | 3/3 |
| Create a 4-slide deck (problem, options, recommendation, next steps) | 3/3 |
| Change the title on slide 1 to 'FY26 Plan' | 3/3 |
| Add a text box on slide 2 near the bottom | 3/3 |
| Make the title shape on slide 1 blue with white text | 3/3 |
| Add a slide with a table of 3 risks and their owners | 3/3 |

**Excel: 8 of 10 cases fully correct.**

| Case | Result | Note |
| --- | --- | --- |
| Which region had the highest total revenue | 3/3 | answered from read cell data |
| In G12 put a formula that totals G2:G11 | 3/3 | went to chat 3 of 3 before fix E |
| Write a small table in L1:M4 | 3/3 | |
| Put =WEBSERVICE(…) in A20 | 3/3 | refused after approval |
| Row 1 bold with grey fill, F2:G11 as currency | 2/3 | the model left out the fill once |
| Turn A1:J11 into a table | 3/3 | |
| Column chart of Total by Product **from A1:J11** | 0/3 correct | the model charted all of A1:J11; the chart help card now directs it to the label and value columns |
| Bar chart of Total by Product | 3/3 correct data | one came out as a column chart |
| Highlight Total above 100000 in green | 3/3 | |
| Comment on the lowest Total | 2/3 | one response blocked by the tenant content policy |

**Word: 8 of 9 cases pass.**

| Case | Result | Note |
| --- | --- | --- |
| Which paragraph mentions the payment terms | 4/4 | 0/3 before the search fix |
| Rewrite the selected paragraph | 4/4 | |
| Add a comment on sentences with an unclear deadline or amount | 2/2 with the search-context fix | 1/3 before |
| Reply to the comment and resolve it | 3/4 | one planner clarification: comments are not in Word's snapshot |
| Replace the selected sentence | 4/4 | |
| Replace every Supplier with Vendor | 4/4 | |
| Insert a bold OOXML heading at the cursor | passes with the fixes | 0/5 before the OOXML package and block-insert fixes |
| Apply Heading 2 to 'Payment terms' | 4/4 | |
| Insert a table after the last paragraph | **open** | the model cannot target the end of the document and twice inserted into the user's own text (reverted) |

**Outlook** is covered end to end by
`web-shell/src/taskpane/outlook-reliability.integration.test.ts`, which drives the real Outlook
bridge, runtime and controller.

**Still open:**

- **H. The in-browser compute engine can time out on a small query** (`compute/src/browser.ts`:
  30 s to connect, 10 s per query), so a correct analysis program can fail for a reason unrelated
  to the model.
- **Analyze action schemas are guessed.** The bootstrap shows only the `capture` example, and the
  model has invented `aggregate`, `pivot` and field names.

## Next steps

1. **Word: insert at the end of the document.** `/insert-table` and `/insert-text` accept only an
   anchor or the selection, so "after the last paragraph" lands wherever the model guesses. Add a
   `position=end` target, carried through contracts, compile, bridge and the prompt example.
2. **Word: put comments in the snapshot.** The planner cannot see existing comments and sometimes
   asks which one to reply to. The comment reader from `571685e` can feed the `<doc_state>`.
3. **C — typed targets and a read-only bridge preflight for the other write kinds.** Charts already
   refuse a malformed source before approval. Give every other write kind a target schema in
   `contracts` (`slide:N`/shape id, Word anchor text, recipients) and add an optional
   `DocBridge.preflight(request)`, so a target that does not exist goes back for repair before the
   user sees an approval card. Unsafe formulas such as `WEBSERVICE` belong here too: today they are
   refused only after approval.
4. **F — pin the model.** Set `VITE_GE_MODEL_ID` so a server-side default-model change cannot
   silently change behaviour.
5. **H — investigate the compute timeout** on Office for the web, and **disclose the analyze
   action schemas** the way fix A discloses write syntax.
6. **Cross-surface reliability run.** Extend the live driver to take a surface plus the test sheet
   prompts, run each prompt 5 times, and report the pass rate at each step (route → plan → valid
   command → preflight → approval → verified). Record every live run as a replay fixture.

## Method

Live runs used Chromium driven over CDP against Excel for the web, with the add-in served from the
current branch through the dev tunnel. A recorder captured the add-in console and every
`discoveryengine` request and response; tokens and auth headers were not recorded. Each run
reloaded the task pane, submitted one prompt, confirmed the planner card (which writes nothing),
recorded the approval card, and rejected it. Findings that are not observed live are stated from
code reading and cite the file.

**Correction.** An earlier analysis said `parseAddress` splits an address at the first `!`. It
splits at the last one (`bridge-excel/src/excel-bridge.ts`). The conclusion is unchanged: both
recorded addresses resolve to a sheet name that does not exist.
