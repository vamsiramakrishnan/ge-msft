# Add-in validation queries

Test fixtures (generated, not committed — regenerate via the scripts noted below if needed):

- `/tmp/addin-validation/sales-schedule.xlsx` — 3 sheets: `Q3 Sales` (deliberately has a blank
  formatting row above its header — a known, **not-yet-fixed** bridge-excel gap from the earlier
  incident investigation, kept here on purpose as a regression probe), `Daily schedule` (matches
  the command-help examples' sheet name exactly), `Budget` (clean, no gotchas).
- `/tmp/addin-validation/project-brief.docx` — headings, a `99.5%` fact for tracked-change
  testing, a 3-row table, risk bullets.
- `/tmp/addin-validation/q3-review.pptx` — 4 slides: title, bulleted risks, bulleted revenue,
  a plain textbox shape.

Open the relevant file in its Office app, sideload the add-in, and paste each numbered query as
one turn — in order, within the same conversation — so later turns exercise multi-turn state
(saved `/work` artifacts, prior turn's selection/results). Sections are independent; you don't
need to run all three surfaces in one sitting, but do run each section's turns in order.

Legend: **✅ expect** = what a healthy response looks like. **⚠️ known gap** = failure here is
expected/documented, not a new bug — note it, don't debug it.

---

## Section A — Excel (`sales-schedule.xlsx`)

1. `ls /doc`
   ✅ expect: lists all three sheets (`Q3 Sales`, `Daily schedule`, `Budget`) plus `outline.md`
   and `selection`, with no cell data pasted into the reply.
2. `find /doc *Sales*`
   ✅ expect: resolves to the `Q3 Sales` sheet entry by name pattern alone.
3. `outline`
   ✅ expect: a short structural summary (sheet names + rough shape), not raw cell dumps.
4. `read 'Q3 Sales'!A1:E12`
   ⚠️ known gap: watch whether row 2 (the blank formatting row) gets treated as the header,
   producing blank/garbled column names or a raw date serial number instead of `2026-07-01` in
   the Date column. If so, that's the documented pre-existing capture-legibility gap — not
   something this turn should silently "fix" by guessing.
5. `save q3.tsv = read 'Q3 Sales'!A1:E12`
   ✅ expect: a named local artifact is created; the reply should be short (an artifact
   handle/summary), not the full pasted range again.
6. `cat q3.tsv head=5`
   ✅ expect: a bounded preview from the SAVED artifact — no new host read should occur for this
   turn.
7. `grep q3.tsv "East"`
   ✅ expect: deterministic line matches for the two East rows, with line numbers — not a
   re-scan of the original range.
8. `ls /work`
   ✅ expect: `q3.tsv` now appears as a workspace entry (multi-turn state check — proves the
   artifact from turn 5 persisted across turns).
9. `read 'Daily schedule'!A1:F5`
   ✅ expect: clean read, headers correct (this sheet has no blank-row gotcha) — a control case
   against turn 4's known gap.
10. `save top-days.md = (read Budget!A1:D5 | sort Variance desc | head 3)`
    ✅ expect: one composed pipeline, one artifact — not three separate turns.
11. `grid Report!A1:B4 = "Region\tRevenue\nEast\t284000\nWest\t474000"`
    ✅ expect: ONE preview/approval card for the whole bulk write (not per-cell `set` calls),
    landing on a fresh `Report` sheet or range.
12. `find /work *.md`
    ✅ expect: `top-days.md` (from turn 10) is found by glob — proves `find`'s glob matching
    works against `/work`, not just `/doc`.

## Section B — Word (`project-brief.docx`)

1. `outline`
   ✅ expect: the heading tree (Executive Summary, Decisions, Risks, Incident Timeline,
   Appendix) — no full-body paste.
2. `search "99.5%"`
   ✅ expect: one exact hit in the Decisions section, with enough anchor context to disambiguate.
3. `suggest "99.5%" => "99.9%"`
   ✅ expect: a **tracked change** (not a silent edit) proposing the replacement, previewed
   before commit.
4. `ls /doc`
   ✅ expect: a document-shaped listing (outline.md, selection, any addressable ranges) — same
   verb, different shape than Excel's sheet listing in Section A turn 1 (proves `ls` genuinely
   adapts per-surface via the same DocFs interface).
5. `find /doc *Risk*`
   ✅ expect: resolves toward the Risks heading/section by name — validates `find` works over a
   Word document's structure, not just Excel sheets.
6. `save risks.md = search "risk"`
   ✅ expect: an artifact capturing the search hits, reusable in later turns without re-searching.
7. `cat risks.md`
   ✅ expect: bounded preview from the artifact.
8. `comment "single-region deployment" "Flag for infra review before sign-off."`
   ✅ expect: an anchored comment on the matched text, not a body edit.

## Section C — PowerPoint (`q3-review.pptx`)

1. `slides`
   ✅ expect: a list of the 4 slides with titles/refs — no full text-frame dump.
2. `outline`
   ✅ expect: deck-level structure (title + slide titles).
3. `ls /doc`
   ✅ expect: a slide-shaped listing — third distinct shape for the same verb, confirming `ls`
   is genuinely surface-agnostic (compare against Section A turn 1 and Section B turn 4).
4. `find /doc *Risk*`
   ✅ expect: resolves to the "Q4 Risks" slide.
5. `slide "Q1 Kickoff" "Hiring plan" "Roadmap review"`
   ✅ expect: one new slide appended, previewed before commit (matches the exact example already
   documented in the skill's capability registry).
6. `shape pp:shape:s4:shape1 "Key risk: multi-region rollout still pending"`
   (Ref may differ — first run `list shape` on slide 4 if the exact `pp:shape:...` id from turn 1
   isn't visible in the reply, then substitute the real id here.)
   ✅ expect: replaces the textbox's text via a captured shape id, not a full-slide rebuild.
7. `save deck-summary.md = outline`
   ✅ expect: an artifact capturing the deck outline, for reuse without re-listing slides.
8. `ls /work`
   ✅ expect: `deck-summary.md` present — cross-surface multi-turn state check (same pattern as
   Section A turn 8).

---

## Cross-cutting checks (run after any section)

- **No-fence resilience**: if any single turn ever comes back with a reply that isn't in the
  expected `cmd` block (rare, but possible on a bad model turn), the add-in should recover
  automatically on the NEXT turn rather than the conversation silently stalling — this is the
  loop-boundary fix from this session. You won't be able to force this deliberately from the UI;
  just note if a conversation ever appears to "hang" after an odd reply instead of self-correcting
  within one extra turn.
- **Bounded context growth**: across a section's full turn sequence, later replies should not
  grow to re-include earlier full pastes (ranges, search results) — `save`/`cat`/`grep` should
  keep replies short even as the conversation gets longer. If a later turn's reply balloons back
  to full-range size, that's worth flagging.

## Regenerating the fixtures

The generator scripts are at `/tmp/addin-validation/gen_excel.py`, `gen_docx.py`, `gen_pptx.py`
(Python, using `openpyxl`/`python-docx`/`python-pptx`). Re-run any of them to regenerate that
file from scratch; they are not part of this repo's committed source.
