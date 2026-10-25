# PowerPoint: adding a slide actually writes its content

**Status:** done, covered by unit tests, and **confirmed live** in PowerPoint for the web ("Add a
slide titled 'Q4 revenue' with the points…" and a 4-slide deck both insert correct slides). The
first version of the fix failed there; the causes were found with live diagnostics and fixed (see
[Round 2](#round-2-what-the-first-fix-missed-on-powerpoint-for-the-web) and
[Round 5](#round-5-a-second-powerpoint-web-quirk)).

> **Later changes (see [POWERPOINT-SHAPE-COMMANDS.md](POWERPOINT-SHAPE-COMMANDS.md)):** the
> slide-adding code is now a shared helper, `appendSlide`, used by `insert-slide` and by the new
> `/add-table-slide slide=new`, which **creates** a slide for its table. `slide=new` is accepted by
> that command only, because there it creates the slide rather than guessing at an existing one.
> Everywhere else it is still rejected. Slide references also accept `pp:slide:<id>`, and the plan
> dependency rules now cover every "last slide" write, not just `add-shape` / `add-table-slide`.

**Scope:** PowerPoint only. Most of the change is in the PowerPoint bridge. There is one small
addition to the plan dependency rules (`runtime/src/planning.ts`), which only affects PowerPoint
write kinds. The command examples the model is shown, and the skill's PowerPoint reference, were
updated to match.

**Files changed:**

| File | What changed |
|---|---|
| `packages/bridge-powerpoint/src/powerpoint-bridge.ts` | new slides get their title and bullets written in, into the layout's boxes or new text boxes; slide references are checked before writing; clear errors instead of "outcome uncertain" |
| `packages/bridge-powerpoint/src/powerpoint-bridge.test.ts` | regression tests; the test fake now models placeholder types, blank layouts, slide size, and PowerPoint web's temporary ids for new slides |
| `packages/runtime/src/planning.ts` (+ test) | in a plan, a `slide=last` write now depends on an earlier "add slide" step, so it's skipped if that step fails |
| `packages/web-shell/src/test-harness/fake-powerpoint.ts` | the end-to-end test fake gives new slides typed title/body placeholders, as real PowerPoint does |
| `packages/runtime/src/command-protocol.ts`, `capability-catalog.ts` (+ test), `packages/contracts/src/command-help.ts` | round 4: the command list the model sees shows `slide` with its syntax and a real example, and ranks it above `/add-table-slide` for "add a slide" (shared code; only PowerPoint's list changes) |
| `packages/contracts/src/capability-registry.ts` | the `/add-shape` and `/add-table-slide` examples use `slide=last` instead of the fake id `s2`/`s3` |
| `skill/m365-surface-commander/scripts/m365-cli-1.0.json` | regenerated from the registry (`bun run skills:generate`) |
| `skill/m365-surface-commander/references/powerpoint-semantics.md` | tells the model which slide references work, and to use `slide "Title" …` for a new slide |

---

## In short

Asking the add-in to add a slide had two problems:

1. **The slide was added but stayed empty.** The add-in reported success, but the title and bullets
   never appeared.
2. **Sometimes nothing happened at all, with a confusing "outcome uncertain" error.** The model had
   tried to put text on a slide that didn't exist yet, the add-in couldn't tell what had happened,
   and the task gave up after three tries.

Both are fixed. A new slide now always gets its title and bullets, and a command that names a slide
the add-in can't find fails with a clear message listing the slides it *can* use.

---

## For everyone: what was going on

### 1. The empty new slide

When you add a slide in PowerPoint, it uses a **layout**: a template that decides which boxes the
slide starts with. "Title and Content" gives you a title box and a bullet box. "Blank" gives you
nothing.

The add-in added the slide, then typed the title into "box number 1" and the bullets into "box
number 2". It never checked whether those boxes existed, or what they were for.

On a deck whose default layout is blank, which includes the test decks we used, there were no boxes
to type into. So nothing was written, but the add-in still said "done".

**Now** the add-in looks at what the new slide actually has:

- If it has a **title box and a bullet box**, the title and bullets go into them. It picks each box
  by what it's *for*, not by its position, so a layout with the boxes in an unusual order still
  works.
- If a box is **missing**, the add-in creates a text box for that part instead. It's sized to fit
  inside the slide, with a large bold title and "•" bullets.
- It writes to the new slide **by its position** in the deck. PowerPoint for the web gives a
  brand-new slide a temporary ID and refuses writes that use it (see Round 2 below). It still
  checks that the slide at that position really is the new one, so if someone else adds a slide at
  the same moment, the text can't end up on their slide.
- If PowerPoint doesn't confirm the write, the add-in says so ("outcome uncertain, check the deck").
  It never says "done" for something it couldn't confirm.

### 2. "Outcome uncertain" when naming a slide

To add a shape or table to a slide, the model has to say *which* slide. Slides in PowerPoint for the
web have internal IDs like `256#0`. The model had written `slide=new`. It wanted "the slide I'm
about to add", but no slide is called that.

The add-in sent "find slide `new`, then add the text" to PowerPoint in one go. When that failed, it
couldn't tell whether any text had landed, so it could only report "outcome uncertain". That's the
safe answer, but it gave the model nothing to fix, so the model gave up.

**Now**:

- The add-in first checks the slide exists, **before** writing anything. If it doesn't, it replies
  with a clear message: *Slide "new" was not found. Use a slide id from the outline (256#0, 257#0,
  …), a slide number 1–9, or "last".* The model can then correct itself.
- It accepts the ways people and models actually refer to slides:

  | You (or the model) write | Means |
  |---|---|
  | `slide=256#0` | that exact slide |
  | `slide=3` | the 3rd slide |
  | `slide=last` | the last slide in the deck, e.g. one just added |
  | `slide=slide:256#0` | same as `256#0` |

  `slide=new` is deliberately **not** accepted as a way to *name* a slide. It sounds like "the slide
  I just made", but if no slide was just made, it would quietly mean "whatever slide is last". It
  gets the "not found, use last" message instead. (The one exception, added later, is
  `/add-table-slide slide=new`, which *creates* a new slide for the table.)
- **If the "add slide" step fails, the next step aimed at "the last slide" doesn't run.** Without
  this, a plan like "add a slide, then put a text box on the last slide" could, when the first step
  failed, put the text box on *your* existing last slide. Now that second step is skipped and
  reported as "skipped, its prerequisite didn't succeed".
- The example commands the model learns from used a made-up slide ID, `s2`. They now use
  `slide=last`, and the model's PowerPoint guide explains the options above.

### Round 2: what the first fix missed on PowerPoint for the web

The first version of this fix worked in the tests but not in PowerPoint for the web. The slide was
added, then the add-in reported "outcome uncertain" and the slide stayed blank. The model then kept
re-reading the deck, trying to work out what had happened.

Two diagnostics run in the browser console found the reason. **When PowerPoint for the web adds a
slide, it first gives it a temporary ID.** It will *read* the slide through that ID, but refuses to
*write* to it, and the ID changes to a permanent one a moment later:

| How the add-in tried to reach the new slide | Result |
|---|---|
| an existing slide, by its ID (control) | ✅ works |
| the new slide, as found in the slide list | ❌ `GeneralException` |
| the new slide, by the ID read right after adding it | ❌ `5010 InvalidParam passed to GetItem(id)` |
| the new slide, **by its position** in the deck | ✅ works |

In that test the new slide's ID was `4123571114#123571113` right after the add, and `267#524649971`
when it was read back later.

So the add-in now writes to the new slide **by position**. It also reads back the slide's permanent
ID afterwards, so later steps (like "add a text box to the last slide") see the real one. And when
a write is uncertain, the message now includes PowerPoint's error code, e.g. *"(PowerPoint:
GeneralException)"*, so it's clear what went wrong.

The tests missed this because the simulated PowerPoint they use didn't have temporary IDs. It does
now, and the version that failed in the browser also fails the tests.

### Round 3: the model kept picking the wrong command

On the next try, the add-in's code was never reached. The model's first command was
`/add-table-slide "Q4 plan" …` (the same mistake as before). It got back "needs a slide", tried
Gemini's own Python tool, and gave up.

Two reasons:

- **The model's PowerPoint guide still told it the wrong thing.** The guide lives in the skill
  bundle loaded into Gemini Enterprise. It is fixed in the repo, but the copy Gemini Enterprise
  uses only changes when the bundle is **republished** (see "Publishing the skill update" below).
- **The error gave the model nothing to go on.** "add-table-slide needs target.slideId" doesn't say
  what to do instead. The errors now say it outright: *"add-table-slide adds a table to an
  EXISTING slide … To create a new slide with a title and bullets, use the slide command instead:
  slide "Title" "bullet" "bullet"."* The "slide not found" error ends with the same hint. So even
  with the old guide, the model is pointed to the right command on its next turn.

### Round 5: a second PowerPoint web quirk

With the model now using `slide "…"`, every attempt still ended "outcome uncertain" and left a blank
slide. A third diagnostic, copying the add-in's exact steps, showed why:

| Step | Result |
|---|---|
| add the slide, find its position by its new ID | ✅ |
| look at the slide at that position (load its ID and shapes) | ✅ |
| write a text box through **that same** slide object | ❌ `5010 InvalidParam passed to GetItem(id)` |
| write a text box through a **fresh** position-based slide object | ✅ |

PowerPoint's error detail showed the reason. **Once a slide object has been loaded, Office.js
quietly switches it to looking the slide up by ID.** For a brand-new slide, that ID is the temporary
one, which rejects writes. The round 2 fix found the slide by position, but it then loaded that same
object to double-check it, which turned it back into an ID lookup.

Now the add-in uses one slide object only to **look** (check it's the right slide, see its boxes),
and writes through **fresh** position-based objects that are never loaded.

### Round 4: why the model picked `/add-table-slide` in the first place

Looking at exactly what the add-in tells the model showed that the model's command list comes from
the **add-in itself**, not from the skill bundle. The add-in builds the list into every request, with
or without a Gemini Enterprise skill. That list steered the model wrong in two ways:

- It showed exact syntax for a fixed set of "core" commands, and `slide` wasn't one of them. The
  model saw `slide` only as a bare name, with no example.
- It also showed the two commands that best matched the request. For "**add** a **slide**…",
  `/add-table-slide` scored highest because both words are in its name, so the model got a full
  `/add-table-slide` card and nothing useful for `slide`.

Now the list shows the exact `slide "Title" "bullet" …` syntax on PowerPoint, and a command whose
**whole name** appears in the request (`slide`) beats one that only shares some words
(`add-table-slide`). The `slide` card also carries a real example:
`slide "Q4 plan" "Hire 5 engineers" "Ship firmware 4.2"`. "Add a table to slide 3" still picks
`/add-table-slide`, because its whole name matches too.

### What you'll notice

- "Add a slide titled X with these points" produces a slide with the title and points on it, on
  any deck, including ones whose default layout is blank.
- On a normal "Title and Content" deck, the text goes into the layout's own boxes, with the
  layout's own bullet style.
- If the model names a slide that doesn't exist, you get a clear "slide not found" message instead
  of "outcome uncertain", and the model usually corrects itself on the next turn.
- If an "add slide" step fails, follow-up steps for "the last slide" are skipped, not applied to
  another slide.

### What it doesn't do (yet)

- It doesn't *choose* a layout. The new slide still uses the deck's default layout. PowerPoint for
  the web refused to tell the add-in which boxes each layout has (it returned `GeneralException`),
  so a reliable "pick Title and Content" isn't possible there yet. The text-box fallback covers the
  gap.
- Text boxes added by the fallback have "•" characters, not PowerPoint's native bullet formatting.

---

## For developers

### Root causes (confirmed on PowerPoint for the web)

A console diagnostic on `Atlas-Mixed-Content-Test.pptx` showed:

```
LAYOUTS:   ['Office Theme › DEFAULT', 'Office Theme › PHOTO_CAPTION']
SLIDE IDS: ['256#0', '257#0', … '264#0']
NEW SLIDE  4123571114#123571113  shapes: []        ← slides.add() on the default layout
```

and the failing command was:

```
/add-shape slide=new type=title text="this is a test slide"   → add-shape — outcome uncertain
```

1. **`insert-slide` (native compose) wrote nothing and reported success.** After `slides.add()`,
   `writeSlideText()` wrote the title into `shapes.items[0]` and the bullets into `shapes.items[1]`.
   Each write was guarded with `if (shape)`, so an empty slide (`shapes: []`) meant no writes, and
   the function still returned `ok: true`. It also assumed position = role, whether or not the
   shapes were placeholders.
2. **`add-shape` / `add-table-slide` batched the slide lookup with the write.** They did
   `slides.getItem(slideId)` plus the add in the **same** `ctx.sync()`, with `mutationQueued = true`
   set before it. An invalid id (`new`) failed the batch, and the catch could only return
   `outcome_unknown`, the correct fail-closed answer for "a mutation may have landed". The model
   received no actionable error, emitted no command on turns 2–3, and the task stopped
   ("exhausted").
3. **The model had no valid way to name the slide it had just added.** It was taught
   `/add-shape slide=s2 …`, with an id no host ever produces.

4. **(Found in round 2.) A just-added slide can't be written through its id on PowerPoint for the
   web.** Right after `slides.add()`, the new slide has a provisional id (`4123571114#123571113`).
   Reads through it work (loading its shapes succeeded), but writes fail: via
   `slides.getItem(provisionalId)` with `5010 "InvalidParam passed to GetItem(id)"`, and via the
   slide object from `slides.items` with `GeneralException`. Writes via `slides.getItemAt(index)`
   succeed. The slide's id then settles (`267#524649971`). The first version of this fix wrote
   through the `items` proxy, so every write failed, and the catch correctly reported
   `outcome_unknown`. The console diagnostics used are `~/Downloads/insert-slide-diagnostic*.js` from
   the investigation.

(`type=title` was not a bug: `shapeTypeFromProp()` in `runtime/src/command-protocol.ts` already maps
unknown types to `textBox`.)

### The changes

#### 1. `insert-slide`: write into the right place, or create one (`powerpoint-bridge.ts`)

```ts
slides.load('items/id');
await ctx.sync();
const existingIds = new Set(slides.items.map((item) => item.id));
slides.add();
try {
  await ctx.sync();
  // Find the new slide's POSITION by the one id that is new (never a stale index)…
  const after = ctx.presentation.slides;
  after.load('items/id');
  await ctx.sync();
  const newIndexes = after.items.flatMap((item, i) => (existingIds.has(item.id) ? [] : [i]));
  const index = newIndexes.length === 1 ? newIndexes[0] : undefined;
  if (index === undefined) return unknownActuationResult(req, '… could not be identified …');
  // …then write BY POSITION: PowerPoint web rejects writes through a new slide's provisional id.
  const provisionalId = after.items[index]?.id;
  const slide = ctx.presentation.slides.getItemAt(index);
  slide.load('id');
  const shapes = slide.shapes;
  shapes.load('items/id,items/type');
  await ctx.sync();
  if (slide.id !== provisionalId) return unknownActuationResult(req, 'The deck changed …');
  await writeComposedSlide(ctx, slide, shapes.items, plan.title, plan.bullets);
  await ctx.sync();
  // ok: true, then — in a fresh request — read the settled id for `location`.
} catch (error) {
  // The slide may already be in the deck; never report an unconfirmed write as success.
  return unknownActuationResult(req, `PowerPoint did not confirm …${hostErrorSuffix(error)}. …`);
}
```

`writeComposedSlide()`:

| Step | Detail |
|---|---|
| find the boxes | `findTextPlaceholders()`: among the new slide's `Placeholder` shapes, on PowerPointApi 1.8 it loads `placeholderFormat.type` (one sync) and picks the title from `Title` / `CenterTitle` / `VerticalTitle` and the body from `Body` / `Content` / `Subtitle` / `VerticalBody`. Below 1.8 it uses the first two placeholders in order (the old convention, now limited to real placeholders). |
| write | the title → the title placeholder; bullets (joined with `\n`) → the body placeholder, which keeps the layout's bullet style |
| fallback | for a part with no placeholder, `slide.shapes.addTextBox()` (PowerPointApi 1.4): title 32 pt bold at 7% from the top; bullets prefixed `• `, 18 pt, from 28% down; 6.7% side margins |
| slide size | `presentation.pageSetup` (PowerPointApi 1.10) when available; otherwise **720 × 405 pt**, the smallest standard slide, so the boxes stay on-slide on 16:9 (960 × 540) and 4:3 (720 × 540) decks too. Only loaded when a fallback box is needed. |
| read vs write (round 5) | a loaded proxy is re-addressed by id (`"statement":"var slide = slides.getItem(...)"` in the host's debug info), so the `getItemAt(index)` probe that is loaded for the id check and the shape list is **only read**. Every write goes through a fresh, never-loaded `slides.getItemAt(index)`, and placeholders through `.shapes.getItemAt(i)` (PowerPointApi 1.3), with `i` taken from the probe's shape order. |
| identify the slide | find its **position** by the single id that wasn't there before the add (not a pre-computed `getItemAt(count)`), then **write by position** (`getItemAt(index)`), because PowerPoint web rejects writes through the new slide's provisional id. Before writing, check the slide at that position still has the new id, so a co-author inserting or deleting a slide meanwhile can't redirect the text. Zero or several new ids, or a mismatch → `outcome_unknown`. |
| settled id | after success, a fresh `PowerPoint.run` reads the slide's id at that position (`readSlideIdAt()`) and reports it in `location`, so provenance and a following `slide=last` see the permanent id. Best effort: if the read fails, the provisional id is kept. |
| confirmation | any failure after `slides.add()` is queued → `outcome_unknown` (`recoveryPending`), never `ok: true`. The message ends with the Office.js error code, e.g. `(PowerPoint: GeneralException)` (`hostErrorSuffix()`; also added to the four other PowerPoint "uncertain" results). |

#### 1b. Errors that point at the right command (round 3)

`add-shape` / `add-table-slide` without a slide now return `no_target` with `SLIDE_TARGET_HINT`
appended:

```
add-table-slide adds a table to an EXISTING slide and needs target.slideId. Pass slide= a slide id
from the outline, a slide number, or last. To create a new slide with a title and bullets, use the
slide command instead: slide "Title" "bullet" "bullet".
```

`resolveSlideRef()`'s "not found" message ends with the same `slide "Title" …` pointer. These
messages reach the model on its next turn: the command loop feeds back the whole `ActuationResult`
(`error.code` and `error.message`) in the ```` ```result ```` block and in
`<execution_state>.observedErrors` (`runtime/src/execution-state.ts`,
`command-context-session.ts`). The loop escapes `<`, `>`, `&` and backticks as `\uXXXX`, so the
hint deliberately uses none of them.

#### 1c. What the model is shown (round 4, shared code, all surfaces)

The per-turn command bootstrap (`runtime/src/command-protocol.ts` `renderCommandBootstrap`) and the
task-ranked command cards (`runtime/src/capability-catalog.ts` `discoverCommands`) are the model's
only command documentation. No Gemini Enterprise skill is needed for them (skills only add
`agentsSpec` plus a mention).

| Change | File | Effect |
|---|---|---|
| `slide` added to the "Common exact signatures" set | `command-protocol.ts` | on PowerPoint, the bootstrap shows `slide "Title" "bullet" ...  OR  slide "Title" (<table expr>)`; the set is filtered by the live manifest, so no other surface changes |
| whole-name bonus (+25) in card ranking | `capability-catalog.ts` | a command whose every name word is in the task outranks one that only shares some words; "add a slide…" → `slide` first, "add a table to slide 3" → `/add-table-slide` first |
| a concrete first example for `slide` | `contracts/src/command-help.ts` | the card's `Example:` is `slide "Q4 plan" "Hire 5 engineers" "Ship firmware 4.2"` (cards skip examples containing `<`); the regenerated `skill/.../m365-cli-1.0.json` picks it up too |

#### 2. Resolve the slide before writing (`resolveSlideRef()`)

Used by `add-shape` and `add-table-slide`. It runs **one read-only sync** (`slides.load('items/id')`)
**before** any mutation is queued:

| Input | Resolves to |
|---|---|
| exact host id (`256#0`) | that slide (checked first, so a numeric-looking real id always wins) |
| `slide:<id>` / `pp:slide:<id>` | the prefix is stripped, then as above |
| `last` (any case) | the last slide. `new` is intentionally not an alias; see section 3. (`/add-table-slide slide=new` is handled separately: it creates a slide.) |
| `1`…`N` | the Nth slide (1-based) |
| anything else | `{ ok: false, degraded: true, error: { code: 'target_conflict', message: 'Slide "<ref>" was not found. Use a slide id from the outline (…up to 12 ids…), a slide number 1–N, or "last".' } }`. The model's `<ref>` is JSON-quoted and cut to 64 characters, so an odd value can't break the message's shape or bloat the next turn. |

After resolution the write is batched as before, with `mutationQueued = true`. So a failure *after*
that point is still reported as `outcome_unknown`. Only a failure that provably happened before any
write is sent is reported as a clean failure. The success `location` now uses the resolved id
(`shape:<real slide id>:<shape id>`), so provenance and undo point at the real slide.

#### 3. `slide=last` depends on a preceding "add slide" (`runtime/src/planning.ts`)

A plan runs its steps in order, and skips a step whose **prerequisite** didn't succeed
(`prerequisite_failed`, which includes an "uncertain" result). Prerequisites are inferred by
`effectResources()` from what each step reads and writes. Before, `insert-slide`, `add-shape` and
`add-table-slide` each had an unrelated opaque key, so nothing linked them. If the insert failed,
`/add-shape slide=last …` still ran, and `last` then resolved to the user's own last slide.

Now:

```ts
case 'insert-slide':
  // Appends a slide, which becomes the deck's LAST slide.
  return { reads: [], writes: [{ kind: 'estate', id: req.kind }, obj('pp:slide:last')] };
case 'add-shape':
case 'add-table-slide':
  return {
    reads: isLastSlideRef(p.target?.slideId) ? [obj('pp:slide:last')] : [],  // /^(slide:)?last$/i
    writes: [{ kind: 'estate', id: req.kind }],
  };
```

A `slide=last` write therefore depends on any earlier `insert-slide` in the same plan. A write to an
exact slide id (`256#0`) or number doesn't. Each kind keeps its existing write key, so ordering
between steps of the same kind is unchanged. This is also why `new` is not accepted: it would
suggest "the slide this plan created" even in a plan that created none.

*Later extended:* `set-shape-text` and `format-shape` addressed at `last` also read it,
`/add-table-slide slide=new` also writes it, and `pp:slide:last` counts as `last` (see
[POWERPOINT-SHAPE-COMMANDS.md](POWERPOINT-SHAPE-COMMANDS.md)).

The security review raised this issue (see [Security review](#security-review)). It would have been
a regression, because before this change any non-id reference simply failed.

#### 4. What the model is taught

- `capability-registry.ts`: the examples are now `/add-shape slide=last shapeType=textBox …` and
  `/add-table-slide slide=last rows=… …`. `m365-cli-1.0.json` is regenerated from it
  (`bun run skills:generate`); `bun run skills:check` confirms it's current.
- `powerpoint-semantics.md`: a new **Slide references** paragraph. It lists the accepted `slide=`
  forms, says not to invent ids like `s2`, `new` or `new-slide`, says a `slide=last` write is skipped
  if the preceding `slide` step fails, and says to use `slide "Title" "bullet" …` alone to create a
  slide with text.
- `powerpoint-semantics.md` also said `/add-table-slide` "creates a slide with structured table
  rows". In round 2 the model believed it, and tried `/add-table-slide "Q4 plan" …` twice
  (`no_target`). It now says the command adds a table to an **existing** slide and needs `slide=`,
  and shows the two-step form: `slide "Title"`, then `/add-table-slide slide=last rows=…`.

`/insert-image` and `/apply-slide-layout` still show `slide=s2`/`s4` in the registry. They aren't
implemented in the bridge, so they're left alone until they are.

### Tests

**Test fakes now behave like real PowerPoint:**

- `bridge-powerpoint` fake: new slides get typed `Title` + `Body` placeholders by default; a deck can
  set `newSlideShapes: []` to simulate a Blank layout; `pageSetup` reports 720 × 405. Its
  requirement check was also fixed: it compared versions as floats, so `'1.10'` read as `1.1` and
  counted as supported on a simulated 1.5 host.
- **(Round 5)** Loading a fake slide proxy now re-addresses it by id, like Office.js, so writing
  through a loaded proxy of a new slide fails too; `ShapeCollection.getItemAt` was added to both
  fakes. The round 2 code fails 5 insert-slide tests against this fake.
- **(Round 2)** The same fake now models provisional ids. A slide from `slides.add()` has a
  temporary id until the current `PowerPoint.run` ends. Any write through it (adding a shape,
  setting text, setting a font) fails the next sync with `InvalidParam passed to GetItem(id)`,
  unless the slide was reached with `getItemAt`. The version of the fix that failed in the browser
  fails 4 tests against this fake.
- `web-shell` end-to-end fake (`fake-powerpoint.ts`): shapes have `type` and `placeholderFormat`,
  and new slides get typed title/body placeholders.

**New tests (`powerpoint-bridge.test.ts`):**

| Test | What it proves |
|---|---|
| matches placeholders by type, not position | a layout with the body box before the title box still gets the title and bullets in the right boxes |
| adds sized text boxes when the default layout has no placeholders | Blank layout → two text boxes (bold title, `•` bullets), both inside the 720 × 405 slide |
| adds only a body text box when the layout has a title but no body | the title goes into the placeholder; only the bullets get a new box |
| falls back to the first two placeholders below PowerPointApi 1.8 | older hosts keep the ordered behaviour |
| writes to the just-added slide by position and reports its settled id | on the provisional-id fake, the title and bullets land, and `location` is the settled id `slide:sim-slide-2` |
| points a missing-slide add-shape / add-table-slide at the slide command | both `no_target` messages name `slide "Title" "bullet"` and the accepted `slide=` forms, and contain no `<` / `>` |
| `capability-catalog.test.ts`: steers "add a slide" to the slide command | on PowerPoint, the bootstrap lists the `slide` signature, the top card is `slide` with the concrete example, and "add a table to slide 3" still ranks `/add-table-slide` first |
| `capability-catalog.test.ts`: lists the slide signature only on PowerPoint | Excel, Word, Outlook, OneNote and Teams bootstraps don't mention `slide "Title"` |
| includes the PowerPoint error code when a write is uncertain | an uncertain result's message contains `(PowerPoint: GeneralException)` |
| never writes onto another slide when the new one can't be identified | a co-author's slide appearing at the same moment → `outcome_unknown`, and their slide is untouched |
| reports an unconfirmed write as outcome-unknown | a failure after the slide was added → `outcome_unknown`, not `ok: true` |
| resolves `slide=last` / `LAST` / `2` / `slide:s1` | each lands the shape on the right slide, and `location` uses the real id |
| fails with the valid slides when the slide doesn't exist | `target_conflict` listing `s1, s2, s3` and suggesting `"last"`; nothing written |
| add-table-slide: resolves `slide=last`, rejects an unknown slide | same behaviour for tables |
| `planning.test.ts`: a `slide=last` write depends on a preceding insert-slide | `add-shape slide=last` and `add-table-slide slide=slide:LAST` depend on the insert, and a failed insert skips both; `add-shape slide=256#0` doesn't depend on it |

Two existing tests ("reports unknown outcome when … landed but its sync failed", for `add-shape`
and `add-table-slide`) made every sync fail. They now let the first, read-only lookup sync succeed,
so they still test what their names say: the write was sent and its confirmation failed.

Against the code before this change, 14 tests fail: 12 of the new ones, plus the 2 updated ones.
The "below PowerPointApi 1.8" test passes on both, because that behaviour was already correct and
is only guarded here. Against the first version of the fix (the one that failed in the browser), 4
insert-slide tests fail on the provisional-id fake. The whole suite is green: `bun run typecheck`,
`bun run lint`, `bun run test` (2,702 passed), `bun run skills:check`, and
`python3 -m unittest test_manifest_contract` in `skill/`.

### Security review

The repo's `security-reviewer` agent reviewed the change, because it touches how writes are
confirmed. It found no critical or high issues, and confirmed that no path reports an unconfirmed
write as success, or reports a clean (retry-safe) failure after a write was sent. Its findings:

| Finding | Severity | Outcome |
|---|---|---|
| A `slide=last` write wasn't tied to the "add slide" step before it, so it could land on the user's own last slide if that step failed; `new` made this worse | medium | **fixed** (section 3; `new` dropped) |
| The new slide was found by index, so a co-author's concurrent slide could receive the text | low | **fixed** (its position is found by its new id and re-checked before writing; round 2 kept this while switching the write itself to by-position) |
| The model's slide reference was echoed unquoted and unbounded in the error | low | **fixed** (JSON-quoted, 64-character limit) |
| The approval card shows `slide=last` / `slide=3`, not the concrete slide it will hit | low | open issue (below) |
| The native `insert-slide` path returns no undo record, though it now knows the slide id | informational, predates this change | open issue (below) |

### Publishing the skill update (optional)

Not required for these fixes (see round 4); it only matters if the add-in is configured with skill
ids.

The changes to `skill/m365-surface-commander/references/powerpoint-semantics.md` and
`skill/m365-surface-commander/scripts/m365-cli-1.0.json` only reach the model once the bundle is
re-uploaded to Gemini Enterprise. Use the runbook in `skill/README.md` ("Dev app runbook"):

```bash
scripts/update-ge-widget-skills.sh --dry-run   # check the target and zips, no changes
scripts/update-ge-widget-skills.sh             # replace the skills with the current bundles
```

It asks for a fresh widget token if needed (a DevTools cURL/HAR from the Gemini Enterprise app) and
writes the new `VITE_GE_*_SKILL` ids into `packages/web-shell/.env`. Rebuild and reload the add-in
afterwards so it uses them.

### How to verify by hand (PowerPoint for the web)

1. **Blank-layout deck:** open `Atlas-Mixed-Content-Test.pptx` and ask *"Add a slide titled 'Q4 plan'
   with the points: hire 5 engineers, ship firmware 4.2"*. Approve. The new last slide should show
   a bold title and two "•" points.
2. **Normal deck:** in a new PowerPoint deck (default "Title and Content"), ask the same. The text
   should go into the layout's own title and content boxes.
3. **Bad slide reference:** ask *"Add a text box saying 'Draft' to slide 99"*. You should see a
   clear "Slide "99" was not found … slide number 1–N" message, not "outcome uncertain".
4. **Last slide:** ask *"Add a text box saying 'Draft' to the last slide"*. It should appear on the
   last slide.
5. **Add a slide, then decorate it:** ask *"Add a slide titled 'Risks', then add a text box saying
   'Draft' to it"*. Both should land on the new slide. If the first step fails, the second should
   show as skipped, not appear on another slide.

---

## Open issues

| Issue | Impact | Plan |
|---|---|---|
| The new slide's layout isn't chosen | slides use the deck's default layout; on Blank-default decks the fallback text boxes are used | pass a layout to `slides.add({ layoutId })` once layout contents can be read reliably on PowerPoint web (`SlideLayout.shapes` returned `GeneralException`) |
| Fallback text boxes use `•` characters | not PowerPoint's native bullet formatting | switch to `paragraphFormat.bulletFormat` once confirmed safe on PowerPoint web |
| The approval card shows the relative reference (`slide=last`, `slide=3`), not the concrete slide | the user approves the command text, not a named slide; the slide actually used is reported afterwards in `location` and can be undone | resolve relative references at preview time, show "slide 7 · '<title>'" on the card, and pass the concrete id to the write |
| Native `insert-slide` returns no undo record | the new slide can't be removed through the add-in's undo | return `inverse: { op: 'delete-object', objectType: 'slide', name: slide.id }` once the undo path supports slides |
| `/insert-image`, `/apply-slide-layout` examples still use `slide=s2`/`s4` | none today (not implemented) | update when those writes are implemented |
