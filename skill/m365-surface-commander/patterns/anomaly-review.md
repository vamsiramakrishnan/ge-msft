# Pattern: anomaly review

A reasoning template, not a command. Read it for shape, then write the turn's algebra.

**Intent:** "find the outliers in <range> and flag them."

**Preconditions**

- the source supports `read`;
- the target supports `comment` (to annotate) and/or `cf` (to highlight).

**Pure core (OBSERVE → DERIVE)**

```
read <range>
  → (derive a threshold purely: e.g. a high quantile, or mean + k·stdev,
     computed with sum/count over the column — never a guessed number)
  → filter rows beyond the threshold
```

Bind the threshold and the outlier set: `let $hi = read … | sum amount` (then reason a cut), and
`let $out = read … | filter amount>$cut`.

**Effect core (EFFECT)**

```
cf      <range> >$cut fill=#FFC7CE      # highlight the whole column at once (one effect), OR
comment <cell> "…"                       # one targeted comment per outlier (anchored on the cell)
```

Prefer **one `cf` rule** over N comments when the request is "highlight" — it is the smaller effect
set (operational law 3). Use `comment` when each outlier needs a distinct, reasoned note.

**Failure rule:** comments/cf anchor on cells you have read; never comment a cell you have not seen.

**Anti-patterns**

- a hard-coded threshold (`>100000`) instead of one derived from the data;
- one comment per row when a single `cf` highlight satisfies the request (effect-budget waste);
- commenting cells outside the range you read.
