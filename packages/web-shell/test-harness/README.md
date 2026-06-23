# Office host simulators + full-stack UI integration harness

The code lives in [`../src/test-harness/`](../src/test-harness/). This README documents it. (The
harness ships *inside* `src/` so it compiles and tests under the same `tsc`/Vitest project as the
rest of `web-shell`; this top-level folder is just its doc home.)

## What this is, and why

The per-surface **bridges** (`@ge/bridge-excel`, `@ge/bridge-word`, `@ge/bridge-powerpoint`) are the
only code in the repo that touches Office.js. They reach the host exclusively through the **global**
entry points the Office runtime installs:

- `Excel.run(cb)` / `Word.run(cb)` / `PowerPoint.run(cb)` — batched object-model access.
- `Office.context.*` — `requirements.isSetSupported`, `document.settings`, `document.customXmlParts`,
  `document.addHandlerAsync/removeHandlerAsync`, `Office.EventType`.

Because that seam is global, we can run the **real, unmodified bridge** against an **in-memory host
simulator**: each `installFake*` writes a richly-typed fake onto `globalThis.{Excel,Word,PowerPoint,
Office}`, seeded with document data. The bridge then reads seeded values and **records its writes
back into the seed**, so a test asserts the post-run host state via `snapshot()` — proving the write
actually landed where the bridge computed it, not merely that the UI said so.

On top of that, [`mountStack`](#mountstack--the-full-stack-wiring) wires the **real** runtime
(`AssistSession`), the **real** `PanelController`, and the **real** React `<App/>` over the fake host
and a scripted model stream. Nothing is mocked except the host object model and the model SSE
stream, so an assertion against the rendered DOM *or* the mutated host snapshot is an assertion about
the whole client stack: bridge → runtime → controller → UI.

The single `unknown` cast that installs fakes onto the typed `globalThis` Office namespaces is
isolated in [`globals.ts`](../src/test-harness/globals.ts) (`installGlobal`); nothing else in the
harness needs a cast or `any`.

## The load/sync model

Every fake follows the same contract, matching how the bridges batch:

- **`load(props)` records which properties to resolve** and returns `this` for chaining.
- **Excel `Range` reads are load-gated** (matching the real host and `office-addin-mock`): a property
  must be named in `load()` **and** resolved by a `context.sync()` before it can be read — reading
  early throws `PropertyNotLoaded`. This keeps an integration test honest: a bridge can't pass by
  reading `.values` without loading it first. Writes (`range.values = …`, `.formulas`, `.format.*`)
  are **not** gated, mirroring Office (you don't load before you write).
- **`sync()` resolves queued loads, then commits queued writes.** For Excel that means promoting
  loaded props to readable and writing back into the seed; for Word/PowerPoint the writes mutate the
  seed object directly, so `sync()` is effectively a `Promise.resolve()`.

This is verified directly (no bridge in the loop) in
[`harness.test.ts`](../src/test-harness/harness.test.ts).

## Per-surface simulator API and fidelity boundary

### Excel — `installFakeExcel(seed?, requirements?)`

Models the slice of the Office.js object model `ExcelBridge` drives:

- `Excel.run(cb)` → a fresh `FakeRequestContext` per run.
- `ctx.workbook.worksheets.getActiveWorksheet()` / `.getItem(name)` (throws on a mis-seeded name).
- `ctx.workbook.getSelectedRange()` → `.address` / `.values`.
- `sheet.getUsedRange()` / `sheet.getUsedRangeOrNullObject()` → `.address` / `.values` /
  `.isNullObject` (an all-blank sheet yields a null object, the ExcelApi 1.4 path).
- `sheet.getRange(a1)` → `.values=` / `.formulas=` (WRITE), `.numberFormat=`,
  `.format.font.bold/italic`, `.format.fill.color`, `.getCell(r,c)` → `.address`,
  `.rowCount` / `.columnCount` / `.isNullObject` (bounded-read metadata).
- `ctx.workbook.names` → `load('items/...')` items `{name,type,formula}`; `.getItemOrNullObject(n)`
  → `.getRange()`.
- `ctx.workbook.comments` → `.add(cellAddress, content)` (WRITE), `load('items/id')` items `{id}`,
  `.items[i].replies.add(text)`, `.items[i].resolved=`, `.onAdded`.
- `sheets.onChanged` / `sheets.onSelectionChanged` / `comments.onAdded` — the `watch()` events.

**Write-record semantics (load-bearing).** Each `FakeRange` the bridge obtains is registered with the
context's `touched` list; on `sync()` every touched range's queued `.values`/`.formulas` commit into
the seed. A `=`-prefixed formula cell takes precedence (recorded verbatim, mirroring Excel routing
`=` cells to `.formulas`); literal cells fall through to `.values`. `getCell(0,0)` resolves the
anchor cell and reads its sheet-qualified `.address` back, so the citation comment lands on the cell
the bridge computed — `applyWriteCells`' write-sync-then-second-sync-to-read-the-anchor ordering is
honored.

A1 helpers (`parseA1`, `addressOf`, `cellRef`, column-letter math) live in
[`a1.ts`](../src/test-harness/a1.ts), kept pure so the fake's address arithmetic is itself testable.

### Word — `installFakeWord(seed?, requirements?)`

Models the seam `OfficeWordHost` drives:

- `Word.run(cb)`; `Word.InsertLocation.replace`; `Word.ChangeTrackingMode.trackAll`.
- `ctx.document.getSelection()` → `.text`.
- `ctx.document.body` → `.text`; `.paragraphs` items `{text, styleBuiltIn}`.
- `ctx.document.body.search(q, {matchCase})` → `items[]` with `.text` and
  `.paragraphs.getFirstOrNullObject()` → `.text` / `.isNullObject` (the context-hint read).
- `range.insertText(text, location)` — the tracked-change WRITE: replaces the first occurrence of
  the anchor in the body and records `{anchor, text}` into `seed.inserts`.
- `range.insertComment(text)` — the comments-as-citations WRITE.
- `ctx.document.changeTrackingMode = trackAll` (recorded for assertions).
- `ctx.document.body.getComments()` → items `{id}`; `comment.reply(text)`; `comment.resolved=`.
- `ctx.document.customXmlParts.add(xml)` — durable-provenance WRITE (via the shared `Office` fake).

**Search re-resolution (load-bearing).** `body.search` is a case-insensitive substring scan, one hit
per containing paragraph. After `insertText(replace)` rewrites the anchor in place, a subsequent
`body.search` for the old anchor returns **zero hits** — so the bridge's drift-degradation path is
exercised, not bypassed. The load/sync ordering (first sync reads hit texts → `chooseAnchorIndex` →
second sync inserts) is preserved end to end.

### PowerPoint — `installFakePowerPoint(seed?, requirements?)`

Models the seam `PowerPointBridge` drives:

- `PowerPoint.run(cb)`.
- `ctx.presentation.getSelectedSlides()` → items `{id, index}`.
- `ctx.presentation.slides` → items `{id, index}`; `.getCount()` → `{value}`; `.add()` (append a
  blank slide, WRITE); `.getItemAt(i)`.
- `slide.shapes` → items; `slide.load('id,index')`.
- `shape.textFrame.textRange` → `.text` (read AND write — the compose-slide WRITE).
- `ctx.presentation.insertSlidesFromBase64(b64)` (prebuilt-deck WRITE; records the payload + appends
  a marker slide).
- Selection/view events go through the shared `Office` bus fake (`addHandlerAsync`).

A freshly `add()`ed slide gets two empty placeholder shapes (title `[0]`, body `[1]`) matching the
bridge's layout convention. `getCount()` is read before `add()`, so `getItemAt(before.value)`
resolves the newly appended slide.

### Office (cross-surface) — `fake-office.ts`

Built fresh per `installFake*`, so each surface gets its own settings bag / handler registry /
requirement set:

- `Office.context.requirements.isSetSupported(name, version)` — the capability gate every bridge
  uses, backed by a seeded `{ name: highestMinor }` map (e.g. `{ ExcelApi: 13 }` ⇒ supports
  `ExcelApi` ≤ 1.13). This drives all the bridges' feature detection — ExcelApi 1.10 for comments,
  1.4 for `getUsedRangeOrNullObject`, 1.7 for named ranges, 1.9 for events; WordApi 1.4 for
  comments/`customXmlParts`, 1.6 for paragraph events.
- `Office.context.document.settings.set/get/remove/saveAsync(...)` — Excel durable provenance.
  Records into `OfficeSeed.settings` and flips `settingsSaved`.
- `Office.context.document.customXmlParts.add(xml)` — Word durable provenance. Appends into
  `OfficeSeed.customXmlParts`.
- `Office.context.document.addHandlerAsync/removeHandlerAsync` + `Office.EventType.
  {DocumentSelectionChanged,ActiveViewChanged}` — host-event registration. Handlers are retained so
  a test fires them via the `OfficeHandlerRegistry` (`officeHandlers.fire(eventType)` / `.count(...)`).

### No-op stubs / fidelity gaps (the boundary)

These host APIs are *called* by a bridge but **stubbed** — a test can pass against the fake that the
real host would compute differently. They are deliberate scope cuts:

- **Excel formula evaluation is not computed.** A `=`-formula write records the formula *string*
  verbatim into the cell; dependent cells are never recomputed. (The composed-read path
  `read | filter | sum` is dry-run by the **runtime** over the seeded values, not by Excel, so those
  tests are unaffected.)
- **`insertSlidesFromBase64` does not parse PPTX.** It records the base64 payload and appends a
  single marker slide; the bridge only needs the call to succeed and the deck to grow.
- **No real slide layouts/masters.** `slides.add()` fabricates two empty placeholder shapes — enough
  structure for the native compose path, not a real master.
- **`load(props)` ignores its property list.** Every property is materialized regardless of what was
  requested, so a bridge that reads a property it forgot to `load()` would *pass* against the fake
  but throw `PropertyNotLoaded` on a real host. (Low risk — the bridges are reviewed for matching
  load/read pairs — but it is a divergence worth knowing.)
- **Excel `numberFormat` / `format.font.*` / `format.fill.color` writes are stored but not committed
  into the cell-value seed** (they sit on the `FakeRange`, which is discarded after the run).
  `snapshot()` exposes only cell `values` and `comments`, so a `format-cells`-only effect is not
  observable via the snapshot — assert formatting through a custom probe if a future test needs it.
- **Word `comment.reply` / `resolved` mutate the seed comment in place**, but there is no requirement-
  set gate on `getComments()` in the fake — the bridge gates its own comment paths via `isSet`, so a
  stale-host (`< WordApi 1.4`) reply test must lower the seeded requirement set to exercise the
  degrade path.

## Seed / fixture shapes

Each surface has a typed seed, a `*Seed(...)` builder that fills defaults, and a rich `default*Seed()`
fixture. Pass a custom seed to `installFake*` (first arg) and/or a requirement-set map (second arg,
e.g. `{ ExcelApi: 9 }` to model a stale host).

### `ExcelSeed` — `excelSeed({ sheets, activeSheet?, selection?, namedRanges?, comments? })`

```ts
interface SheetSeed      { name: string; origin: string; values: string[][] }  // origin = top-left A1
interface NamedRangeSeed { name: string; range: string }                       // sheet-qualified A1, no '='
interface CommentSeed    { id: string; cell: string; content: string; replies: string[]; resolved: boolean }
interface ExcelSeed {
  sheets: SheetSeed[]; activeSheet: string; selection: string;                 // sheet-qualified A1
  namedRanges: NamedRangeSeed[]; comments: CommentSeed[];
}
```

`excelSeed` defaults `activeSheet`/`selection` to the first sheet's name/origin and `namedRanges`/
`comments` to empty. `defaultExcelSeed()` is an FSI workbook: a `Sales` sheet (header + six regional
revenue/cost rows), a `Summary` sheet, a `SalesTable` named range, and a starting selection over a
data row.

### `WordSeed` — `wordSeed({ selectionText?, paragraphs, comments? })`

```ts
interface WordParagraphSeed { text: string; styleBuiltIn: string }             // style → heading level
interface WordCommentSeed   { id: string; text: string; replies: string[]; resolved: boolean }
interface WordSeed {
  selectionText: string; paragraphs: WordParagraphSeed[]; comments: WordCommentSeed[];
  changeTrackingMode?: string;                                                 // recorded by the bridge
  inserts: Array<{ anchor: string; text: string }>;                           // recorded tracked-change writes
  addedComments: Array<{ anchor: string; text: string }>;                     // recorded insertComment writes
}
```

`wordSeed` defaults selection/comments to empty and initializes the `inserts`/`addedComments` write
logs. `defaultWordSeed()` is a contract-ish MSA: a heading + body paragraphs including an SLA claim
to redline, plus one existing comment.

### `PowerPointSeed` — `powerPointSeed({ slides, selectedIndices? })`

```ts
interface ShapeSeed { text: string }
interface SlideSeed { id: string; shapes: ShapeSeed[] }                        // shapes[0] = title, by convention
interface PowerPointSeed {
  slides: SlideSeed[]; selectedIndices: number[];                             // zero-based; bridge reads items[0]
  insertedDecks: string[];                                                    // recorded base64 merges
}
```

`powerPointSeed` defaults `selectedIndices` to `[0]` (or `[]` for an empty deck) and `insertedDecks`
to empty. `defaultPowerPointSeed()` is a 3-slide Q3 business review (title + two content slides),
with slide index 1 selected.

## `mountStack` — the full-stack wiring

```ts
const sim = installFakeExcel();                       // 1. install the fake host (caller owns restore)
const ui  = mountStack({                              // 2. wire the real stack over it
  surface: 'excel',
  client: scriptedClient([ /* model turns */ ]),
  triggers,                                            //    optional actuation gate / event registry
});
```

`mountStack` installs nothing itself — the caller must `installFake*` first. It then:

1. picks the **real** bridge via `selectBridge(surface)`,
2. constructs a **real** `AssistSession(bridge, client, { unit, triggers? })`,
3. constructs a **real** `PanelController(session, bridge)`,
4. renders the **real** `<App/>` with `createRoot` + `act`.

It returns the live `{ bridge, session, controller, container, root }` plus:

- **`flush()`** — drains a fixed batch of microtasks inside `act()` (for synchronous-ish settling,
  e.g. after `refreshContext`).
- **`waitFor(predicate, timeoutMs=1000)`** — pumps the event loop (real `setTimeout(0)` macrotasks)
  inside `act()` until `predicate(controller.getState())` holds. **Why:** the command loop
  *suspends on async stream consumption between gates* — it `await`s the model stream and the
  approval promise — so a test must wait for a gate to be *staged* (`pendingPlan !== undefined`) or
  the run to *settle* (`!busy`) rather than assume a fixed number of microtasks.
- **`act(fn)`** — runs a discrete user gesture (`approvePlan`, `rejectPlan`, `runCommands`,
  `refreshContext`) inside `act()`, draining the immediate microtasks it schedules so the rendered
  DOM stays warning-free.
- **`unmount()`** — unmounts React and removes the container. It does **not** restore host globals;
  the simulator owns that, so `afterEach` should call `ui?.unmount()` then `sim?.restore()`.

The **scripted client** (`scriptedClient(turns)`) is a fake `StreamAssistClient`. Each turn is a
string (or `{ text, citations }`) carrying a ` ```cmd ` block; the loop pulls the next turn per
`runCommands`/`ask` and replays it as `token` → `citation*` → `provenance` → `done` SSE events, so the
real streamed-message / citation / provenance-stamped-write paths all run. It seeds a default
provenance identity (`sim.user@acme`), which is what an approved write persists into the host. The
returned `{ client, queries }` lets a test assert what was sent to the model; `mountStack` accepts
either the wrapper or a bare `client`.

## Adding a new surface simulator

1. **Enumerate the exact host calls its bridge makes.** Read the bridge (and any host port) and list
   every `Host.run` object-model call and `Office.context.*` call it touches — that list *is* the
   fidelity boundary, and belongs in the fake's file header.
2. **Model only that slice** as small fake classes with `load()` no-ops and immediate property
   materialization. Don't implement the full `@types/office-js` surface.
3. **Install the global** via `installGlobal(name, fake)` (composed with the shared `Office` fake
   from `makeFakeOffice`), and return a `restore()` from `composeRestores`.
4. **Record writes into the seed** at the same point the bridge syncs — directly on the seed object
   (Word/PowerPoint) or via a `touched`-then-commit-on-`sync()` list (Excel, when writes are queued
   on a proxy and read back). Preserve any read-then-write sync ordering the bridge depends on.
5. **Expose `snapshot()`** returning a deep-copied, assertion-friendly view of the post-run host.
6. **Add a fidelity test** in `harness.test.ts` driving the global directly (no bridge) to prove the
   load/sync, write-record, and event semantics — so an integration test can't pass against a fake
   that diverges from the host. Export the new symbols from `index.ts`.

## Writing a new integration test

Pattern (jsdom; start the file with `// @vitest-environment jsdom`):

```ts
sim = installFakeExcel(/* optional seed, optional requirements */);
ui  = mountStack({ surface: 'excel', client: scriptedClient([
  '```cmd\nset Summary!B2 42\n```',
  '```cmd\ndone\n```',
]) });
await ui.flush();

let run!: Promise<void>;
await ui.act(() => { run = ui.controller.runCommands('write a value'); });
await ui.waitFor((s) => s.pendingPlan !== undefined);   // wait for the staged gate

await ui.act(() => ui.controller.approvePlan());        // the user gesture
await ui.waitFor((s) => !s.busy);                       // wait for the run to settle
await run;
await ui.flush();

// Assert the MUTATED HOST, not just UI state:
expect(sim.snapshot().sheets.find((s) => s.name === 'Summary')?.values[1]?.[1]).toBe('42');
```

Always assert the **host snapshot** for write outcomes (so the test fails if a write leaks through or
fails to land), and remember the teardown:

```ts
afterEach(() => { ui?.unmount(); sim?.restore(); ui = undefined; sim = undefined; });
```

To exercise a **stale host**, lower the requirement set: `installFakeExcel(undefined, { ExcelApi: 9 })`
makes comments (ExcelApi 1.10) unsupported, so the bridge degrades instead of throwing. To exercise
the **fail-closed gate**, pass a `TriggerRegistry` with a `pre-actuation` handler returning
`{ kind: 'block' }` and assert the host snapshot is unchanged.
```
