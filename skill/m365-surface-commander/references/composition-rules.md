# Composition rules

How to assemble a correct program from the [algebra](algebra.md). The type laws say what _composes_;
these say how to _build well_.

## Operational laws

1. **Derive before mutating.** Do all reading and pure transformation first; emit effects last. A
   program reads, computes, then writes — never interleaves a write between two derivations.
2. **Bind once, reuse.** If a value is used twice, `let $v = …` it and reuse `$v`. Never repeat a
   `read` of the same source — a second read can drift from the first.
3. **Emit the smallest effect set that satisfies the request.** Don't write three cells when one
   formula does it; don't chart raw source when the request is "top N" (chart the derived range).
4. **Never manufacture an address.** Selectors, object names, ranges, comment ids — every one must
   come from something you read or the user gave you. A guessed `Sales!F2` or `{comment-id}` is wrong.
5. **An effect result is not a value.** You cannot read back what a `spill`/`set`/`table` produced in
   the same program. Compute the value _before_ the effect; compute a dependent range from the
   table's known size (see §derived ranges), not by "reading the effect."
6. **No hidden scratch.** Never stash intermediate data in spare worksheet cells or document text as
   scratch memory. Bindings (`$v`) are your working memory; the document is the deliverable.
7. **Only the injected signature.** Use only the verbs/operators in this turn's `<capabilities>`. If
   it isn't in the signature, it does not exist this turn — do not reach for another app's verb.

## Derived ranges (the spill → table/chart pattern)

A `spill` writes a `Table` into a grid starting at an origin cell. Its size is known at plan time
(rows × columns of the bound table), so the **resulting range is computable** — `spill Report!A1`
of a 10×2 table occupies `Report!A1:B11`. Reference that computed range in the dependent `table` /
`chart` / `cf`. You never bind the spill's result; you compute its range. The dependent effects then
depend on the spill (the preflight's `plan` shows this grouping).

## The composition decision procedure

For each turn, in order:

1. Inspect this turn's `<capabilities>` signature — what verbs/operators exist now.
2. Decide if a **direct command** is enough. If so, emit it; stop here.
3. **Read the smallest source** that answers the request — a range, not the whole sheet.
4. **Bind** any value you'll reuse (`let $v = …`).
5. Perform **all** transformation **purely** (filter/select/sort/head/aggregate).
6. Compute the **concrete effect targets** (the exact ranges/anchors).
7. Emit the **minimum dependent effect set**.
8. If the program crosses the complexity threshold, **preflight it** (`surface_cli`, below).
9. Output **exactly one** `cmd` program.

## Canonical normal form

Structure every non-trivial program in four phases, in this order:

```
OBSERVE   read / search / outline            — get the data
DERIVE    let bindings + pure transforms     — compute purely
EFFECT    bounded, concrete host mutations   — the smallest effect set
VERIFY    (optional) read-after-write later  — confirm, in a *later* turn
```

This is more useful than "keep it under N lines." A program that reads, derives, then emits a tight
effect set is correct by construction; one that interleaves is where bugs live.

## When to break a program (semantic boundaries, not line count)

Keep a coherent OBSERVE→DERIVE→EFFECT chain in **one** program. Start a **new** program/turn only
when one of these is true:

- **A fresh observation is required.** You write, the host recalculates, and you need the _new_
  values (e.g. write formulas → read the computed results → chart them). The new values don't exist
  until the first effect lands — that's a second OBSERVE phase, a new turn.
- **A host-minted id is consumed downstream.** An effect produces a non-deterministic id you must
  then target. That's an effect-result dependency, not pure composition.
- **Approval authority changes.** An in-document edit and an externally-visible send (mail/post)
  should not ride one approval just because they're in one script.
- **The artifact or surface changes.** Excel analysis → PowerPoint creation is a handoff, not one
  transaction across two add-in instances.
- **Failure domains differ.** A reversible formatting change and an irreversible external post should
  not be one atomic unit.
- **An effect budget is exceeded.** A plan that would mint hundreds of comments must be partitioned
  or rejected before approval, even if it is algebraically valid.

## Preflight with `surface_cli`

Run the deterministic checker before emitting when the program is non-trivial (more than one binding,
more than two effects, any dependent materialization, or near a limit) — it catches unknown verbs,
out-of-signature verbs, unbound `$vars`, budget overruns, and wrong dependencies that are easy to get
wrong by hand. See the invocation policy in [SKILL.md](../SKILL.md) (Preflight section). The runtime
parser remains authoritative; the preflight only catches structural errors earlier.

## Anti-patterns

- **Sort after head** — `… | head 10 | sort …` ranks the wrong 10 rows. Sort, then head.
- **Chart the source for a "top N"** — chart the derived/spilled range, not the raw range.
- **Re-reading instead of binding** — two reads of one source drift; bind once.
- **Hidden scratch cells** — never; use `$` bindings.
- **A verb from another app** — only this turn's signature exists.
