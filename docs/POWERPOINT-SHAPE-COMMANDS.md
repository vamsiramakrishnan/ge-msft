# PowerPoint: editing shapes, adding text boxes, formatting, and table slides

**Status:** done. Test cases 48–51 were each run live in PowerPoint for the web (browser automation
on the test deck `Atlas-Mixed-Content-Test.pptx`) and pass. Covered by unit tests.

**Scope:** mostly PowerPoint. Two changes are in shared code and apply to every app: the routing
of edit requests (48) and the ranking of the commands the model is shown (48–51). Both are written
so that only PowerPoint's behaviour changes in practice. Test 47 was also checked; see
[the end of this doc](#test-47-compiled-multi-slide-deck).

**Files changed:**

| File | What changed |
|---|---|
| `packages/web-shell/src/taskpane/components/App.tsx` | "change / set / rename / replace / format / recolour / move …" requests go to the editing flow instead of plain chat |
| `packages/runtime/src/capability-catalog.ts` (+ test) | the model is shown the right command for "… on slide N" requests and for colours |
| `packages/runtime/src/command-protocol.ts` (+ test) | commands accept the forms models actually write (bare slide refs, `color=`, `fill.color=`, quoted table rows) |
| `packages/runtime/src/planning.ts` (+ test) | a "last slide" write can't land on the wrong slide if the step that added a slide failed |
| `packages/bridge-powerpoint/src/powerpoint-bridge.ts` (+ test) | shapes and slides are found by id, number, `last` or `title`; off-slide positions are rejected; `/add-table-slide slide=new` creates the slide; clearer errors |
| `packages/contracts/src/command-help.ts`, `capability-registry.ts`, `command-grammar.ts` | what the model is told about `shape`, `slide`, `/add-shape`, `/format-shape`, `/add-table-slide` |
| `skill/m365-surface-commander/…` (`powerpoint-semantics.md`, regenerated `m365-cli-1.0.json`) | the same guidance for the skill bundle |

---

## In short

| # | Test | Before | After |
|---|---|---|---|
| 48 | Change the title on slide 1 to 'FY26 Plan' | ❌ went to plain chat; Gemini *claimed* it changed the title, nothing changed | ✅ title changed, nothing else touched |
| 49 | Add a text box on slide 2 near the bottom | ❌ wrong slide, unsupported slide reference, position below the slide; nothing added | ✅ text box on slide 2, inside the slide |
| 50 | Make the title shape on slide 1 blue with white text | ❌ the model couldn't name the title shape and was never shown `/format-shape`; nothing changed | ✅ blue fill, white text, only that shape |
| 51 | Add a slide with a table of 3 risks and their owners | ⚠️ table added to the **existing** last slide, on top of its content | ✅ new slide with a title and a native table below it |

---

## For everyone: what was going on

The add-in lets Gemini edit the deck through a small set of commands. The model has to pick the
right command and fill it in correctly. In these four tests, three kinds of thing went wrong.

### 1. Some edit requests never reached the editing flow (test 48)

The add-in only sends a typed request to the editing flow if it **starts with an action word it
recognises**: *add, create, insert, update, rewrite…*. "**Change** the title…" wasn't on the list,
so it went to plain chat. Gemini then said "I have changed the title", but no command ran, and the
deck was unchanged. (The add-in's own count said "0 changes applied".)

**Fix:** *change, set, rename, retitle, replace, format, recolour, move, resize* and *highlight* now
count as edit words too. This only decides which flow handles the request. Every edit still shows
you a plan and waits for **Approve** before anything changes.

### 2. The model wasn't shown the right command (tests 48, 49, 50)

For each request, the add-in shows the model its full command list, plus detailed "cards" for the
two commands that best match the request. The matching was fooled in two ways:

- **"on slide 1" counted as asking for the `slide` command.** So every "…on slide N…" request showed
  the model the *add a slide* command, not the shape or format command it needed.
- **Colour words matched nothing.** "Make it blue with white text" didn't point at `/format-shape`,
  which is the only command that changes colours.

**Fix:** "slide 1", "slides 2–4" and similar phrases are treated as *where* to act, not *what* to
do, and colour names count as "colour". The descriptions of `/format-shape`, `/add-table-slide` and
`slide` were also sharpened. For example, `slide` now says "for a native table, use
`/add-table-slide`".

**Looking at a slide didn't show its shapes (test 50).** When the model read a slide to find
the title, it got the slide's text but no shape ids, so it had nothing to point a command at. In
one take it kept re-reading the slide and never made the change. **Fix:** reading one slide
(`read slide:1`) now also lists that slide's shapes. Each line shows the shape's id and type
(text box, picture…), and the title is marked. Reads also accept a slide's own id, not just
its number.

### 3. Commands rejected what the model naturally wrote (tests 49, 50, 51)

Once the model reached the right command, it often wrote it in a slightly different form from the
one the add-in accepted, and the command failed. The add-in now accepts the forms seen live:

| Model wrote | Now means |
|---|---|
| `/add-shape pp:slide:257#0 …` or `/add-shape 2 …` | that slide (by id or number) |
| `shape pp:shape:1:title "FY26 Plan"` | slide 1's **title** shape |
| `/format-shape … color=white` | **text** colour white (`fill=` is the background) |
| `/format-shape … fill.color=#0000FF font.color=#FFFFFF` | same as `fill=… fontColor=…` |
| `/add-table-slide "Risks" "Risk \| Owner" "Delay \| Pat"` | a **new** slide titled "Risks" with that table |
| `/add-table-slide slide=new title="Risks" rows=…` | a **new** slide with that table |

When something still doesn't fit, the add-in now **explains what to do** instead of just failing,
and the model fixes it on the next turn. That happened in every test:

- a shape it couldn't find → *"Shape "1" was not found on that slide (there are no shape numbers). For the slide title use title as the shape: shape pp:shape:1:title "…" (or shape=title on /format-shape). Otherwise use one of its shape ids: …"*
- a position below the slide → *"That position is outside the slide, which is 720 × 405 pt …"*
- text into a picture → *"Shape id 2 is a Image and has no text. For the slide title use shape pp:shape:1:title "…", or the id of a text box."*
- `/format-shape` with no formatting → rejected, instead of reporting "applied" for a change that
  did nothing
- a text box with no text, or a shape with no position → rejected, with the slide's size and a
  "near the bottom" example. (Seen while recording: a malformed `/add-shape slide:2 {"text": …,
  "position": "bottom"}` had been "applied" as an empty box in the top-left corner.)

### What you'll notice

- "Change / rename / recolour / move …" requests produce a plan to approve, not a chat reply.
- Titles can be changed and formatted by just saying "the title on slide N".
- Text boxes land on the slide you named, inside the slide.
- "Add a slide with a table of …" makes a new slide with a real PowerPoint table under its title.
- The model sometimes still needs one extra turn (e.g. it tries `pp:shape:1:1` first, is told "use
  shape=title", then succeeds). The final result is right, but the activity summary may still say
  "did not complete all its operations", because it counts the rejected first attempt.

---

## For developers

### Live evidence (before)

| # | Command the model emitted | Result |
|---|---|---|
| 48 | *(none: routed to chat)* | Gemini text: "I have changed the title…"; `0 changes applied` |
| 49 | `/add-shape pp:slide:264#0 type="textbox" … left=100 top=500 …` | `add-shape — no_target` (slide 9, not 2; positional ref not parsed; top 500 on a 405 pt slide) |
| 50 | `/format-shape $title.refId fill.color="#0000FF"` …, then `list shapes \| where role == "title"` … | `format-shape — no_target` ×2; `/format-shape` never appeared in the command cards |
| 51 | `/add-table-slide slide=last rows=…` | applied, but onto the existing last slide (overlapping its bullets); no new slide |

### The changes

#### Routing (`App.tsx`, shared)

`OFFICE_ACTION_REQUEST_RE`, consulted only by `shouldUsePlannerForFreeText`, gained
`change|set|rename|retitle|replace|format|recolou?r|colou?r|move|resize|highlight`. It only chooses
the planner over chat for free text; writes still go through `approvePlan` / `approveWrite`.

#### Command cards (`capability-catalog.ts`, shared)

`discoverCommands` ranks commands per task:

- `slide N` / `slides N–M` phrases are removed before scoring. For the whole-name bonus they still
  count, but only if some name word also appears outside them, so "add a table to slide 3" still
  ranks `/add-table-slide` first.
- `COLOR_WORDS` (red, blue, green, …, white, black, grey) normalise to `color`.
- Registry and help text: `/format-shape` `useWhen` mentions colour and "make a title blue with
  white text"; `/add-table-slide` mentions `slide=new`; `slide` says to use `/add-table-slide` for
  tables. The `shape` signature is `shape <pp:shape:slide:shape> "text"`, with a hint showing
  `pp:shape:1:title`.

Checked in `capability-catalog.test.ts`: the top 2 cards for the four test prompts include
`/format-shape`, `shape`, `/add-shape` and `/add-table-slide` respectively.

#### Parsing (`command-protocol.ts`)

- `withPositionalSlide`: for `add-shape` / `add-table-slide`, a first positional token shaped like
  a slide reference (`pp:slide:…`, `slide:…`, `N`, `N#M`, `last`) becomes `target.slideId` when no
  `slide=` was given. Quoted titles never match.
- `format-shape`: `dottedStyleProps` maps `fill.color=` / `line.color=` / `font.color|size|bold…=`
  (the tokenizer only accepts `\w[\w-]*` keys, so these arrive as positional tokens); `fillColor=`,
  `lineColor=`; `color=` / `textColor=` → the font colour.
- `add-table-slide`: `title=` → `params.slide.title`; `positionalTable` turns `"Title" "A | B" "1 | 2"
  …` (every row with the same ≥ 2 cells, split on `|` or tab) into `slide=new` + title + rows. A flat
  list of single cells is left alone, since its columns can't be inferred.

#### Bridge (`powerpoint-bridge.ts`)

| Helper / change | What it does |
|---|---|
| `resolveShapeTarget(ctx, slideRef, shapeRef)` | used by `set-shape-text` and `format-shape`, read-only before any write. Slide via `resolveSlideRef` (id, `pp:slide:` / `slide:` prefix, 1-based number, `last`). Shape by **exact id** or **`title`**: the Title/CenterTitle placeholder (API 1.8), else the first shape with text. **No shape numbers:** shape ids are small numbers themselves (2, 3, 4…), so `2` would be ambiguous. An unknown ref returns the slide's shape ids and types. |
| `set-shape-text` | loads the shape `type`; a non-text shape (image, chart, table) → *"…is a Image and has no text…"* before any write |
| `format-shape` | an empty `shapeFormat` → `no_format` (it used to report `ok: true` for doing nothing) |
| `missingGeometryMessage` | `add-shape` without `left`/`top`/`width`/`height` fails before writing. The message gives the slide size and a "near the bottom" example; the host would otherwise drop the shape at the top-left corner. A `textBox` with no `text` fails too (`no_text`). |
| `offSlideMessage` + `knownSlideSize` | `add-shape` / `add-table-slide` with explicit geometry outside the slide (`pageSetup`, PowerPointApi 1.10) fail before writing, with the slide size in the message |
| `appendSlide(ctx)` | shared by `insert-slide` and `add-table-slide slide=new`: appends, finds the new slide's position by its new id, checks it through a loaded probe, and leaves writes to fresh `getItemAt` proxies (PowerPoint web rejects writes through a new slide's provisional id; see [POWERPOINT-INSERT-SLIDE.md](POWERPOINT-INSERT-SLIDE.md), rounds 2 and 5) |
| `add-table-slide slide=new` | creates the slide, writes an optional title (`writeComposedSlide`: title placeholder or text box), and places the table **below the title** (27% down, inside the margins) unless the command gave a position; `location` uses the settled slide id; the undo record deletes the **slide** |
| `add-table-slide` with `title=` on an existing slide | rejected (it used to be silently ignored) |
| `readRange` (`read slide:N`, and `inspect slide:N`, which falls back to it) | accepts a host slide id (`256#`, `slide:256#`, `pp:slide:256#`) and `last`, not only a slide number (`slideIndexForSelector`). A single-slide read also appends a **shape listing** (`slideShapeListing`): one line per shape with its id, type, and text, a `title` marker on the shape that `shape=title` resolves to, and a header saying which `slide=` / `shape=` values to use. Before this, a read returned only the slide's text, so the model could not find a shape id to target. On one recording of test 50 it ran `inspect slides`, `inspect slide:1` and `inspect outline` over and over and never wrote anything. Attached-context chips are unchanged; only reads get the listing. |

#### Plan dependencies (`planning.ts`)

`effectResources` models "the deck's last slide" as a resource:

- `insert-slide` and `add-table-slide slide=new` **write** it.
- `add-shape`, `add-table-slide`, `set-shape-text` and `format-shape` addressed at `last` /
  `slide:last` / `pp:slide:last` **read** it.

So if a plan's "add slide" step fails or is uncertain, the follow-up `…slide=last` write is skipped
(`prerequisite_failed`) instead of landing on the user's own last slide.

### Security review

The repo's `security-reviewer` reviewed the targeting changes. No critical findings. Outcomes:

| Finding | Severity | Outcome |
|---|---|---|
| New relative-slide paths (`set-shape-text` / `format-shape` on `last`, `add-table-slide slide=new`, `pp:slide:last`) weren't tied to a preceding "add slide" in the plan | high | **fixed** (`planning.ts`, with tests) |
| Numeric shape refs are ambiguous (id vs position) | medium | **fixed**: shape numbers removed; exact id or `title` only |
| `slide=new` undo would delete only the table | medium | **fixed**: undo deletes the created slide |
| `title=` silently ignored on an existing slide | low | **fixed**: rejected |
| Relative / `title` refs resolve at apply time, not at approval (the card shows `slide=last` / `shape=title`, not the concrete shape) | medium | **open issue** (below) |
| Settled slide id read by position without re-checking; shape undo records name no slide | low | **open issue** |
| Routing regex / echoing refs in errors | none | no finding: routing only chooses planner vs chat; echoed refs are JSON-quoted, capped, and listings contain ids and types only |

### Tests

New tests (plus two existing sync-count tests updated for the extra read-only lookups):

| File | Covers |
|---|---|
| `powerpoint-bridge.test.ts` | title by `slide 1 / title` (only it changes); format via `title`; unknown shape → shape-id list; a shape number is not a position; no text into a picture; empty format rejected; `pp:slide:` accepted; off-slide position rejected with the size; `slide=new` creates a titled slide with the table below the title (placeholder and blank layouts) and records the slide as its undo; `title=` on an existing slide rejected |
| `command-protocol.test.ts` | positional slide refs; quoted titles never taken as slides; `slide=new title=`; quoted `"A \| B"` rows → a new titled table slide; flat lists not guessed; `fill.color=` / `font.color=`; `color=` → text colour; `fillColor=` |
| `planning.test.ts` | `set-shape-text` / `format-shape` / `add-shape pp:slide:last` depend on a preceding insert; `add-table-slide slide=new` counts as adding the last slide |
| `capture.test.ts` | slide selectors (ids, `pp:slide:`, numbers, `last`, names rejected before any host call); the shape listing format (title marker, one line per shape, text clipped to 80 characters) |
| `powerpoint-bridge.test.ts` (reads) | `read` by slide id and `pp:slide:` ref; a single-slide read lists each shape's id and type and marks the title |
| `capability-catalog.test.ts` | the top command cards for the four test prompts |
| `command-surface.integration.test.ts` | "change / set / rename / replace / format / move" route to the planner; a question doesn't |

The whole suite is green: `bun run typecheck`, `bun run lint`, `bun run test` (2,726 passed),
`bun run skills:check`, and `python3 -m unittest test_manifest_contract`.

### Live results (after)

| # | What the model did | Final state |
|---|---|---|
| 48 | `shape pp:shape:1:1 …` → "use shape=title" → `shape pp:shape:1:title "FY26 Plan"` | slide 1 title = "FY26 Plan"; logo, subtitle and footer unchanged |
| 49 | `/add-shape slide=2 … top=400` → "outside the slide, which is 720 × 405" → `… top=340` | "Draft – not for distribution" text box on slide 2 at top 340 |
| 50 | `/format-shape pp:shape:1:1 …` → "use shape=title" → `/format-shape pp:shape:1:title fill="blue" color="white"`. Re-run with the shape listing: `inspect slide:1` → `/format-shape pp:shape:1:1 …` (rejected, no shape 1) → `/format-shape pp:shape:1:3 fill=#0000FF text=#FFFFFF` (id 3 from the listing) | title fill `#0000FF`, text `#FFFFFF`; logo, subtitle and footer unchanged |
| 51 | `/add-table-slide "Risks and Mitigations" "Risk \| Owner \| Mitigation" "…" "…" "…"` (one command) | new slide 16: title, native 4 × 3 table below it; "1 changes applied; 0 not applied" |

---

## Test 47: compiled multi-slide deck

"Create a 4-slide deck: problem, options, recommendation, next steps for moving to a 4-day week"
**inserts 4 correct slides** (checked live). But it does it as four ordinary `slide "Title"
"bullet" …` commands, not the **compiled-deck** path (`/insert-slide` with a staged `.pptx`) the
test case expects. That path can't run in this build: `@ge/deck-compiler` and
`buildPowerPointDeckImportRequest` exist but are **not called anywhere in the app**, only by their
own unit tests. "Build a decision brief" is just a prompt template that goes through the same
planner. So the outcome passes, and the expected *mechanism* needs either correcting in the test
case or wiring in as a feature.

---

## A second deck: titles without a title placeholder

**Status:** fixed and confirmed live. On "Tokenomics LLM Evaluation Framework" (40 slides, a Google
Slides export), "change the title to Q4 plan on the first slide instead of what is tokenomics"
now changes slide 1's title and nothing else.

### For everyone

The test deck worked, but on this deck the same request changed nothing. Three things were wrong:

1. **The add-in listed shapes by position.** It told the model about "Shape 1" and "Shape 2" on the
   slide. Here "Shape 1" is a thin decorative line above the title. The model aimed at "shape 1",
   which is either no shape at all or the line.
2. **"The title" meant "the first shape with text."** Slides made in Google Slides have no real
   title box. On most of this deck's slides the first text is a small label such as
   "TOKENOMICS / 01", which sits above the actual title.
3. **When a shape wasn't found, the error listed shape numbers and types only**, so the model
   could not tell the line from the title. On its second try it picked the line.

**Fixes:** shapes are now listed by their real id ("Shape 86 on slide 1"), and empty shapes such
as lines aren't listed at all. "Title" now means the text set in the largest font; if two are the
same size, the one nearer the top wins. The "not found" error now shows each shape's text and
marks the title.

**Still worth knowing:** with **Whole deck** attached (40 slides), Gemini Enterprise switched to
its Python code-execution tool. It then tried to run the add-in's commands as Python, which fails
every time, and gave up after 41 and 60 steps. With Whole deck removed, the same request
worked in one pass. A single-slide edit doesn't need the whole deck attached.

### For developers

| Change | File | What it does |
|---|---|---|
| Shape refs titled by id | `capture.ts` `shapeContextRef` | `Shape <shapeId> on slide <n>`, not `Shape <ordinal>`; the `ordinal` parameter is gone |
| Only text shapes listed | `powerpoint-bridge.ts` `listContext` | skips shapes with no text: they resolve to no context and were write targets for text |
| Title fallback by font size | `powerpoint-bridge.ts` `findTitleShape` → `largestTextShape` | without a Title/CenterTitle placeholder, pick the text shape whose first visible character (`textRange.getSubstring(i, 1).font.size`) is largest, ties → smallest `top`. The first character because a range with mixed runs reports `font.size: null`. One read-only sync; falls back to the first shape with text if the host reports no sizes |
| Descriptive not-found error | `powerpoint-bridge.ts` `describeShapes` | `id 85 (GeometricShape, no text), id 86 (TextBox, title: "What is Tokenomics?")`; shape text is quoted, single-line, clipped to 40 characters |

Measured on the deck: title placeholders none; per-slide first-character sizes 12 pt (section
label) vs 28.5 pt (title) on slides 2, 3, 5 and 6. On slide 4, the largest text is
"Now what do we optimize?" (34.5 pt).

Live runs, same prompt:

| Run | Context | What happened |
|---|---|---|
| Before the fix | Whole deck | `shape pp:shape:1:1` (not found) → code execution → `shape pp:shape:1:title` + `shape pp:shape:1:85` (the line); plan rejected, nothing written |
| After the fix | Whole deck | code execution from turn 1 (`Python code execution requested` / `failed`), 60 steps, nothing written |
| After the fix | none | `inspect "Slide 1"` → `read` → `shape pp:shape:1:86 "Q4 plan"`: "1 changes applied; 0 not applied"; shape 85 unchanged |

Tests (`powerpoint-bridge.test.ts`): `listContext` skips empty shapes and titles by id; title by
largest font over a smaller label above it; font-size tie → topmost; the not-found listing shows
text, `no text` and the title marker. The fake host gained `Shape.top` and
`TextRange.getSubstring().font.size`.

---

## Open issues

| Issue | Impact | Plan |
|---|---|---|
| The approval card shows relative refs (`slide=last`, `shape=title`, `slide=2`), which resolve at apply time | the user approves the command text, not a concrete shape; the resolved target appears afterwards in `location` | resolve refs during the dry run and show "slide 1 · title (id 3)" on the card; apply exact ids only |
| Stale Microsoft sign-in token: `WIF token exchange failed … stale to sign-in` | the add-in fails and doesn't recover on its own (seen during testing) | on this error, fetch a fresh Entra token and retry the exchange once |
| Compiled-deck insert not wired in (test 47) | multi-slide decks insert as N separate slides | wire `buildPowerPointDeckImportRequest` into the plan flow, or correct the test case |
| With a large attachment (Whole deck, 40 slides), Gemini Enterprise may switch to its Python code-execution tool and try to run the command block as Python | the task loops until it gives up; nothing is written. The runtime's re-prompt ("Hosted Python/code execution is not a valid executor response") does not stop it | engine-side: turn off code execution for the assistant used by the add-in, or check whether a request option can disable it. Client-side: don't resend the whole-deck attachment on command turns for a single-slide edit |
| The model sometimes guesses shape ids (`pp:shape:1:0` … `1:10`) in its reasoning before reading the slide | none seen: only the read and the final correct command ran. But on a deck with small shape ids a guess could hit a real shape | the approval card should name the concrete shape (see the first row) |
| The model often needs one corrective turn (e.g. `pp:shape:1:1` → `title`) | the activity summary still counts the rejected attempt | show resolved examples in the core signatures; the error hints already make the fix one turn |
