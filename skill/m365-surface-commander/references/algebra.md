---
title: Value Algebra
kind: reference
skill: m365-surface-commander
topics: [values, pipelines, transforms, effects, type-laws]
load_when: A command needs pure computation, table shaping, aggregation, or composed write values.
---

# The value algebra

This is the language under the commands. The commands you emit are a surface syntax over a small,
typed algebra: **reads produce values, pure operators transform values, and effects consume a value
and terminate.** Learn the shapes here once; the exact operators available **this turn** come from
the injected `<capabilities>` signature (machine-readable in `scripts/m365-cli-1.0.json`) — this file
teaches the _types and laws_, the signature is _authoritative_.

## Values

| Type       | What it is                                          | Where it comes from                           |
| ---------- | --------------------------------------------------- | --------------------------------------------- |
| `Table`    | rows × named columns                                | `read` of a range; any pure transform         |
| `Number`   | one scalar number                                   | `sum`/`avg`/`min`/`max`/`count` of a column   |
| `Text`     | one string                                          | a cell/selection read; a literal              |
| `Boolean`  | true/false                                          | a predicate                                   |
| `Selector` | an address that names where to read (`Sales!A1:G9`) | you write it; never invent one                |
| `RangeRef` | a concrete written range (`Report!A1:B11`)          | computed from a `spill`'s origin × table size |

You **cannot see content until you read it.** A `Selector` is a promise of where data is; a `Table`
is the data once read. Never fabricate a value, a selector, an object id, or a comment id.

## Operators

### Pure — compose freely, never terminate

```
read    : Selector            -> Table        # the only way to obtain content
filter  : Table × Predicate   -> Table        # rows where the predicate holds
select  : Table × ColumnList  -> Table        # keep these columns
sort    : Table × Column × Dir -> Table       # order by a column (asc|desc)
head    : Table × Number      -> Table        # first N rows
tail    : Table × Number      -> Table        # last N rows
sum     : Table × Column      -> Number        # also avg / min / max
count   : Table               -> Number        # row count
```

A pipeline is left-to-right composition of these: `read … | filter … | sort … | head N`. The output
type is the **last** operator's result type — a trailing `sum`/`count` yields a `Number`, otherwise a
`Table`.

### Effect — consume a value, produce a reviewable change, TERMINATE

```
set     : Range × Scalar  -> Effect          # one cell
grid    : Range × Grid    -> Effect          # a literal rectangular cell grid
spill   : Range × Table   -> Effect          # a whole table as a grid (the table→grid sink)
table   : Range           -> Effect          # promote a range to a native Table
chart   : Range × Spec    -> Effect          # a chart over a range
cf      : Range × Rule    -> Effect          # a conditional-format rule
suggest : Anchor × Text   -> Effect          # Word tracked change
comment : Anchor × Text   -> Effect          # a comment
reply   : CommentRef × Text -> Effect        # reply to a comment
slide / page / mail / post / compose -> Effect
```

**The critical distinction:** a pure operator returns a value you can keep composing. An **effect
terminates computation** — you cannot pipe out of it, and its result is _not_ a value you can read
back in the same program. (Effects that mint a host id you must then target — e.g. a posted message
you later edit — are a separate, deferred case; see ADR-0008 §3. For everything in this turn's
signature, effects are terminal.)

## Type laws (what composes, what doesn't)

- `filter(p, select(C, T))` is valid **only when** the columns `p` tests are a subset of `C` —
  selecting away a column then filtering on it is a type error.
- `head(n, sort(k, T)) ≠ sort(k, head(n, T))` — **sort before you take.** Taking the first N then
  sorting gives you the wrong N rows. This is the single most common composition bug.
- `sum(k, T) : Number` and `select(C, T) : Table` — an aggregation collapses a Table to a scalar;
  a projection keeps it a Table. A `set` wants a scalar; a `spill` wants a Table.
- `set(range, value) : Effect`, `grid(range, cells) : Effect`, and `spill(range, table) : Effect`
  all terminate. You bind the _value_ before the effect (`let $x = …`), never the _effect_.
- Use `grid` for a literal rectangular payload. Use `spill` for a `Table` value produced by a
  pipeline. Do not expand either into many scalar `set` effects.

## How the algebra reaches the commands

- A **read** command (`read`/`search`/`outline`) or a `let $v = <pipeline>` binding produces a value.
- A bound value is reused by name (`$v`) — **bind once, reuse**, never re-read the same source.
- An **effect** command consumes concrete values/addresses you have already computed.

For the operational laws (derive-before-mutate, smallest-effect-set, the decision procedure, the
OBSERVE→DERIVE→EFFECT→VERIFY normal form, and when to break a program), load
[composition-rules.md](composition-rules.md). For exact per-app syntax and limits, load
[capability-map.md](capability-map.md).
