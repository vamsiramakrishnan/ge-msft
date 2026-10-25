# PowerPoint: "Whole deck" attach — speed, mixed content, tables

**Status:** done and tested locally. PowerPoint web confirmed the speed and mixed-content fixes on
real decks. Table reading is covered by unit tests and still needs a check in PowerPoint. On branch
`fix/powerpoint-whole-deck-attach`.

**Scope:** PowerPoint only. All changes are in the PowerPoint bridge.

**Files changed:**

| File | What changed |
|---|---|
| `packages/bridge-powerpoint/src/powerpoint-bridge.ts` | batched slide reader, safe handling of pictures/charts/placeholders, table cells, per-shape fallback |
| `packages/bridge-powerpoint/src/capture.ts` | `tableValuesToText()` — turns a table grid into text lines |
| `*.test.ts` next to each of the above | regression tests |

---

## In short

There were three problems, all now fixed for PowerPoint:

1. **"Whole deck" looked like it did nothing.** It was actually attaching, but it took about
   **16 seconds** for a 9-slide deck on PowerPoint for the web, with no spinner. It now takes about
   **1.5 seconds**, roughly **11× faster**. Every question you ask in PowerPoint also starts faster,
   because the add-in reads the deck the same way before each question.
2. **A picture could break the whole deck.** One picture, chart or table anywhere (including a picture
   dropped into a layout's content box) made the read fail. Those shapes are now handled safely.
3. **Tables were invisible.** Numbers that only appeared in a PowerPoint table were never sent to
   Gemini. Table cells are now read.

---

## For everyone: what was going on

### 1. Speed

Think of the add-in and PowerPoint as two people passing notes through a slow mail slot.

- **Before:** for each slide, the add-in sent one note asking "what shapes are on this slide?",
  waited for the answer, then sent another asking "what text is in those shapes?", and waited again.
  A 9-slide deck took 19 trips through the slot.
- **After:** the add-in asks about every slide in one note, then asks for all the content in
  another. It's 3–4 trips, whatever the size of the deck.

Each trip takes about 0.85 seconds on PowerPoint for the web. On the desktop apps a trip is much
faster, which is why the problem mainly showed up on the web.

| | Trips to PowerPoint (9 slides) | Time measured on PowerPoint web |
|---|---|---|
| Before | 19 | **16.3 s** |
| After | 3 | **1.45 s** |

### 2. Pictures, charts and tables

Asking PowerPoint "what text is in this picture?" makes it refuse, and one refusal cancelled the
whole batch of questions. The add-in now checks what each shape is first and only asks text
questions of shapes that hold text. One tricky case: a picture placed in a layout's content box
still says it's a "placeholder", so the add-in also checks what the placeholder contains.

If PowerPoint still refuses a shape unexpectedly, the add-in falls back to asking one shape at a
time and skips the ones it can't read. The rest of the deck still attaches.

### 3. Tables

PowerPoint can hand over a whole table as a grid of cells in one go. The add-in now asks for that
grid in the same trip as the text, so tables add no extra wait. Each row becomes one line, for
example `Madrid | $310K | 10 November 2026`.

### What you'll notice

- **Whole deck** attaches in a second or two.
- Decks with pictures, charts, tables, lines, callouts and placeholder pictures attach fine.
- Questions about numbers in a **table** get answered from the table.

### What it doesn't do (yet)

- Words drawn inside a **picture** and values in a **chart** are not read. Office.js has no way to
  read them.
- **Speaker notes** are not read.
- Only the **first 60 slides** are read, the same limit as the add-in's other deck reads.
- Tables need PowerPoint **API 1.8**. PowerPoint for the web and current Microsoft 365 desktop apps
  have it; on older versions tables are skipped.
- The chip still shows no loading indicator while it attaches.

---

## For developers

### Root causes

**Speed.** `resolveContext()` for the `pp:deck` ref looped over every slide and awaited
`readSlide()` for each one. `readSlide()` made two `ctx.sync()` calls: one to load the shape
collection, one to load each shape's `textFrame.textRange.text`. That's `1 + 2 × N` round-trips for
N slides. On PowerPoint for the web each `ctx.sync()` is a full network round-trip (about 850 ms
measured). The controller only marks the chip attached once `attachRef()` resolves and shows no
pending state, so from the outside it looked like a silent failure. The same loop sat behind
`readAllSlides()`, which feeds `captureDocState()` (the `<doc_state>` snapshot sent every turn) and
`searchDocument()`.

We measured this in the PowerPoint web console on a 9-slide deck:

```
per-slide loop  : 19 syncs → 16,280 ms
batched         :  3 syncs →  1,448 ms
```

**Mixed content.** `Shape.textFrame` **throws `InvalidArgument`** for shapes without a text frame,
and in a batched sync one throw fails the whole batch. A per-shape diagnostic on PowerPoint web (45
shapes, `Atlas-Mixed-Content-Test.pptx`) showed exactly which shapes refuse:

| `shape.type` | `placeholderFormat` (`type/containedType`) | Text frame? |
|---|---|---|
| `TextBox`, `GeometricShape`, `Callout` | — | ✅ yes (including empty decorative shapes) |
| `Line` | — | ✅ yes (always empty) |
| `Placeholder` | `Title/null`, `Body/null` | ✅ yes |
| `Image`, `Table`, `Chart` | — | ❌ throws `InvalidArgument` |
| **`Placeholder`** | **`Content/Image`** | ❌ **throws `InvalidArgument`** |

A picture in a layout placeholder still reports `type: 'Placeholder'`; only
`placeholderFormat.containedType` reveals the image.

### The changes

#### 1. One batched reader: `readSlides(ctx, slides)` (`powerpoint-bridge.ts`)

It uses at most three syncs, whatever the slide count (four for the whole deck, counting the slide
list):

```ts
// sync 1: every slide's id/index + every shape's id/type
const shapeCollections = slides.map((slide) => {
  slide.load('id,index');
  const shapes = slide.shapes;
  shapes.load('items/id,items/type');
  return shapes;
});
await ctx.sync();

const candidates = shapeCollections.flatMap((shapes) => shapes.items.filter(isReadableShape));
const readable = await withoutNonTextPlaceholders(ctx, candidates); // sync 2, only if placeholders
const texts = await readShapeContent(ctx, readable);                 // sync 3: text + tables together
```

| Step | Helper | Syncs | When |
|---|---|---|---|
| shape ids + types | inline | 1 | always |
| drop placeholders holding a picture/table/chart | `withoutNonTextPlaceholders()` | 1 | only if the batch has placeholders and the host has PowerPointApi 1.8 |
| text frames + table cells | `readShapeContent()` | 1 | always (per-shape fallback on failure) |

Every read path now goes through it:

| Caller | Before | After |
|---|---|---|
| `resolveContext` (`pp:deck`, Whole deck) | its own per-slide loop, **no slide cap** | `readAllSlides()` → capped at `MAX_READ_SLIDES` (60) |
| `readAllSlides()` → `captureDocState`, `searchDocument` | per-slide loop | `readSlides()` |
| `readSlide()` (selection, single slide, `read slide:N`, chip listing) | its own 2-sync read | wrapper over `readSlides(ctx, [slide])` |
| `readShapeContext()` (a single shape chip) | always asked for `textFrame` → threw on a picture/table chip | same type/placeholder checks; table chips return their cells, picture chips return nothing |

The `SlideElement` shape is unchanged, so `slidesToContext`, `slideElementsToDocStateBlocks` and
`searchSlides` needed no change.

#### 2. Which shapes are read (`powerpoint-bridge.ts`)

```ts
const TEXT_SHAPE_TYPES = new Set(['GeometricShape', 'TextBox', 'Placeholder', 'Callout', 'Freeform']);

function isReadableShape(shape: PowerPoint.Shape): boolean {
  if (shape.type === 'Table') return isSet('PowerPointApi', '1.8');
  return TEXT_SHAPE_TYPES.has(shape.type);
}
```

Placeholders get one more check (`placeholderHoldsText`): a `containedType` of `null` (empty or
text placeholder) or a text type is read; `Image`, `Table`, `Chart` and the rest are skipped.

Unread shapes (`Image`, `Chart`, `Group`, `Media`, `SmartArt`, …) stay in the slide's shape list
with empty text, so reveal and addressing still work.

#### 3. Tables (`powerpoint-bridge.ts` + `capture.ts`)

`queueShapeRead(shape)` queues the right read for each shape and returns a getter for after the
sync:

```ts
if (shape.type === 'Table') {
  const table = shape.getTable();
  table.load('values');                                 // string[][], PowerPointApi 1.8
  return () => tableValuesToText(table.values ?? []);
}
const range = shape.textFrame.textRange;
range.load('text');
return () => range.text ?? '';
```

`tableValuesToText()` (in `capture.ts`, pure) produces one line per row, joining cells with ` | `.
It collapses whitespace inside a cell, since a row must stay one line, and drops rows that are
entirely blank. The result flows through `shapesToSlideText()` like any other shape text, so it
lands in the slide's body lines, the `<doc_state>` outline and deck search.

#### 4. Per-shape fallback (`readShapeContent()`)

It reads every candidate in one sync. If that sync fails (a shape the host refuses unexpectedly, or
placeholders on a host older than 1.8), it retries **one shape per sync** and skips any that throw.
This is slower, but it only runs for decks that contain something unexpected, and the rest of the
deck still attaches. The web diagnostic showed that the request context stays usable after a
refused shape, which is what makes this retry safe.

### Tests

**The PowerPoint fake** (`powerpoint-bridge.test.ts`) now behaves like the real host:

- shapes have a `type` (default `'TextBox'`); placeholders have a `containedType`;
  `placeholderFormat` and `getTable()` / `Table.values` exist;
- touching a text frame, placeholder format or table a shape doesn't have **fails the next
  `ctx.sync()`**, i.e. the whole batch, exactly as PowerPoint does.

**New tests:**

| File | Test | What it proves |
|---|---|---|
| `powerpoint-bridge.test.ts` | reads the whole deck in a constant number of syncs | 3 and 30 slides use the **same** number of syncs (at most 3) |
| | skips shapes without a text frame | a deck with an `Image` and a `Table` still attaches |
| | skips a placeholder holding a picture | text placeholders are read, the picture one is skipped, in exactly 4 syncs (no fallback) |
| | falls back to per-shape reads | a shape refused unexpectedly is skipped; the rest of the slide is read |
| | still reads the deck without `placeholderFormat` | on PowerPointApi 1.5 the picture placeholder is handled by the fallback |
| | reads native table cells into the whole deck | `Site \| Budget`, `Madrid \| $310K` are read, still in 3 syncs (tables share the text round-trip) |
| | skips tables without `Table.values` | on PowerPointApi 1.5 tables are skipped, the rest still attaches |
| | resolves a table shape chip / picture shape chip | a table chip returns its cells; a picture chip returns `[]` instead of throwing |
| | bounds the whole-deck read to `MAX_READ_SLIDES` | slide 61+ is not read |
| `capture.test.ts` | flattens a table grid | one line per row, whitespace collapsed, blank rows dropped |

Each regression test was checked against the code before its fix and failed there. The full suite
is green: `bun run typecheck`, `bun run lint`, `bun run test` (2,681 passed).

### How to verify by hand (PowerPoint for the web)

1. **Speed:** open a text-only deck of about 10 slides (e.g. `Atlas-Q3-Review-TextOnly.pptx`) and
   click **Whole deck**. It should turn attached (`◈`) in about 2 s.
2. **Mixed content:** open `Atlas-Mixed-Content-Test.pptx` and attach **Whole deck**. It should
   attach in a couple of seconds.
3. **Tables:** with that deck attached, ask "What is Madrid's budget?" → **$310K** (slide 5 table),
   and "How many picks per hour does aisle C-14 get?" → **312** (slide 9 table).
4. **Still out of reach (expected):** "How many dock doors does Hamburg have?" (only inside a
   picture) and "What was Milan's route-length reduction?" (only in the chart).

---

## Open issues

| Issue | Impact | Plan |
|---|---|---|
| No pending state on a chip while it attaches | slow reads look like nothing happened | show a loading state in the context tray (`web-shell`) |
| Speaker notes aren't read | notes facts aren't available to Gemini | add a notes read (the whole-deck code comment used to claim it) |
| Picture and chart content isn't read | facts only inside pictures/charts aren't available | no Office.js text API; would need image/chart export plus a separate extraction step |
