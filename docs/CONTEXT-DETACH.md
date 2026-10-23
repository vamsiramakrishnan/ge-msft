# Context tray: removing a chip really removes its content

**Status:** done and covered by unit tests. Not yet checked in a live Office app. On branch
`fix/powerpoint-whole-deck-attach`.

**Scope:** every surface: PowerPoint, Word, Excel, Outlook, OneNote and Teams. The fix is in the
shared session code, not in any one app's bridge.

**Files changed:**

| File | What changed |
|---|---|
| `packages/runtime/src/assist-session.ts` | tracks which content each chip attached; removing a chip removes all of it, except content another attached chip still needs |
| `packages/runtime/src/assist-session.test.ts` | regression tests |

---

## In short

When you removed a chip from the context tray (the `×` on **Whole deck**, **Slide**, **Whole
document**, …), the chip disappeared, but **its content often stayed attached and kept being sent to
Gemini with your next questions**. So answers could still come from something you had just removed.

Now, removing a chip removes everything that chip attached. When two chips share the same
content, that content stays attached until you remove the last of them.

---

## For everyone: what was going on

Before sending a document to Gemini, the add-in cuts it into smaller pieces ("chunks"), so a long
deck or document fits comfortably. Each piece gets its own label.

Take the **Whole deck** chip, labelled `pp:deck`. Its pieces were labelled `pp:deck#0`,
`pp:deck#1`, `pp:deck#2`, and so on. When you removed the chip, the add-in threw away anything
labelled exactly `pp:deck`, which matched none of the pieces. So every piece stayed attached.

It's like asking someone to remove "the Q3 report" from a pile. They look for a folder with exactly
that name and find nothing, because it was filed as "Q3 report, part 1", "part 2" and "part 3". So
everything stays in the pile.

The same thing happened to any chip whose content was big enough to be cut up. For example:

| App | Chips affected |
|---|---|
| PowerPoint | **Whole deck**, **Slide** |
| Word | **Whole document**, **Paragraph** |
| Excel | ranges, sheets and tables |
| Outlook, OneNote, Teams | the email, page or transcript chips |

### What changed

The add-in now remembers which pieces came from which chip. Removing a chip removes all of its
pieces.

### When two chips share the same content

Sometimes two chips point at the same content:

- In **Excel**, a "Selection" chip and a "Range" chip can cover the same cells.
- In **Outlook**, every chip reads the open email.
- In **OneNote**, every chip reads the open page.
- In **Teams**, every chip reads the captured meeting transcript.

If you've attached both chips and remove one, the add-in keeps the pieces the other chip still
needs. The content only goes away once no attached chip uses it. That way, removing one chip never
silently takes content away from another.

### What you'll notice

- Removing a chip stops its content from being used in your next question.
- Re-attaching a chip after the content changed (e.g. you deleted slides or paragraphs) no longer
  leaves the old, deleted parts behind.
- With two overlapping chips attached, removing one of them keeps the shared content. Removing
  both removes it.

---

## For developers

### Root cause

- `SessionContext` (`packages/gemini-client/src/session-context.ts`) stores each attached
  `ResolvedContext` in a `Map` keyed by **its own** `ref.id`.
- `@ge/content` chunks content with ids `${sourceId}#${index}` (`packages/content/src/chunk.ts`).
  So `resolveContext(ref)` returns items whose ids differ from the chip's ref id: `pp:deck` →
  `pp:deck#0…n`, `word:document` → `word:document#0…n`, `pp:slide:<id>` → `pp:slide:<id>#0…n`.
- `AssistSession.detach(id)` only did `this.context.remove(id)`. Nothing is stored under the chip
  id, so every chunk stayed attached and kept going out as `queryParts`.
- Re-attaching had a related gap. Chunks that no longer existed (a deck that shrank from 5 chunks
  to 3) were never removed, because the new resolution simply overwrote ids `#0–#2` and left
  `#3–#4`.

A single item whose id happens to equal the chip id (e.g. a PowerPoint shape chip) was unaffected.
That's why the bug only showed up on chunked content.

### The change (`assist-session.ts`)

`AssistSession` now records which resolved ids each chip attached:

```ts
/** refId → the resolved-context ids that ref added. */
private readonly attachedParts = new Map<string, Set<string>>();
```

**Attach** (`attachRef()` and the auto-attach path `attachContext()`) resolves the ref, releases the
ref's previous parts (the re-attach case), then adds and records each new part:

```ts
const resolvedParts = await this.toolOperation('context:resolve', { ref }, () =>
  this.bridge.resolveContext(ref),
);
if (this.attachmentVersions.get(ref.id) !== version) return; // detached while resolving
this.releaseAttachedParts(ref.id);
for (const resolved of resolvedParts) this.addAttachedPart(ref.id, resolved);
```

**Detach** releases the ref's parts, and also removes the ref id itself unless another chip claims
it:

```ts
detach(id: string): void {
  this.attachmentVersions.set(id, (this.attachmentVersions.get(id) ?? 0) + 1);
  this.releaseAttachedParts(id);
  if (!this.isClaimedByAttachedRef(id)) this.context.remove(id);
}
```

**Shared parts are only removed when no other attached chip claims them:**

```ts
private releaseAttachedParts(refId: string): void {
  const parts = this.attachedParts.get(refId);
  if (!parts) return;
  this.attachedParts.delete(refId);
  for (const partId of parts) {
    if (!this.isClaimedByAttachedRef(partId)) this.context.remove(partId);
  }
}
```

### Why the shared-part guard matters

Several bridges resolve **different chips to the same ids**:

| Surface | Chips that share resolved ids | Why |
|---|---|---|
| Excel | a **selection** chip and a **range** chip over the same cells | ids derive from the cell address |
| Outlook | every chip | `resolveContext` ignores the ref and reads the one active mail item |
| OneNote | every chip | always reads the active page |
| Teams | every chip | always reads the captured transcript window |
| Word, PowerPoint | none | each chip's chunks have distinct ids |

Before, the shared item was simply overwritten in `SessionContext`, and detaching removed nothing,
so this never came up. Once detach started removing parts, a naive version would have removed
content a still-attached chip needs. The guard prevents that. The tests cover it with two Outlook
chips over one mail item.

### Behaviour details

- **Detach during an in-flight attach:** `attachmentVersions` still guards this. If the chip is
  removed while its content is being read, the late result is dropped and nothing is added.
- **Content added outside the chip tray** (the per-turn `<doc_state>` snapshot, the working-context
  brief, lazy `read`/`search` results) doesn't go through `attachRef()`/`attachContext()`. It isn't
  tracked here, and its existing lifecycle is unchanged.
- **Context compaction** can evict a chunk on its own. The chip's record may still list the evicted
  id. That's harmless: removing a missing id is a no-op, and the stale record is dropped when the
  chip is removed or re-attached.
- **No bridge or contract changes.** `ContextRef`, `ResolvedContext`, `SessionContext` and every
  `resolveContext()` are untouched. `web-shell`'s `PanelController.detach()` already calls
  `session.detach(chipId)`, so the context tray needed no change.

### Tests (`assist-session.test.ts`)

| Test | What it proves |
|---|---|
| lets the caller attach/detach context explicitly *(existing)* | a single unchunked item (id = chip id) still detaches |
| detaches every chunk a PowerPoint ref resolved to | attach → 3 chunks; re-attach after the deck shrinks → 2 (stale chunk gone); detach → 0; the auto-attach path records and releases the same way |
| detaches every chunk on any surface | a Word **Whole document** ref's 2 chunks are both removed on detach |
| keeps a part another attached ref still claims | two Outlook chips share one mail item: detaching the first keeps it, detaching the second removes it |

Each new test fails against the code before the fix. The full suite is green: `bun run typecheck`,
`bun run lint`, `bun run test` (2,684 passed).

### How to verify by hand

1. **PowerPoint:** attach **Whole deck**, then remove it (`×`). Ask something only the deck
   answers. It should no longer be answered from the deck.
2. **Word:** attach **Whole document**, remove it, and ask about something only in the document. It
   should no longer be answered from the document.
3. **Excel (overlapping chips):** select some cells, then attach both the selection chip and a range
   chip over the same cells. Remove one: questions about those cells should still be answered.
   Remove the other: they should no longer be answered from the sheet.
4. **Re-attach after an edit:** attach **Whole deck**, delete a slide, attach **Whole deck** again.
   The deleted slide's text should no longer be used.

---

## Open issues

| Issue | Impact | Plan |
|---|---|---|
| Not yet checked in a live Office app | behaviour is covered by unit tests only | run the hand checks above in PowerPoint, Word and Excel |
