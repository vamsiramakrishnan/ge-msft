---
title: Example Session - Excel
kind: example
skill: m365-surface-commander
surface: excel
topics: [range, formula, comment, composition]
load_when: A concrete Excel command-loop transcript is needed.
---

# Example session — Excel

A multi-turn session: read a region, compute a per-region total, write it back as a native
formula, and flag an anomaly with a comment. Shows read-before-write, composition, and the
one-block-per-turn protocol.

**User task:** "Add a March total for the East region in F2, and flag any day with a spike."

---

**Turn 1 — orient and read.**

````
**thought** I need the structure, then the East region's March values.
```cmd
outline
read Sales!A1:F20
```
````

The add-in replies with a `result` block containing the outline and the range values.

---

**Turn 2 — compute and write back as a formula, then flag the spike.**

After reading, the East rows are `Sales!C2:C8`, and `C6` is ~5x its neighbors.

````
**thought** Write a native SUMIF so it stays live, then comment the outlier.
```cmd
set Sales!F2 =SUMIF(A2:A8,"East",C2:C8)
comment Sales!C6 "Anomalous spike — ~5x the surrounding days; verify before reporting."
```
````

Both writes are previewed as one change set, approved once, then applied and recorded.

---

**Turn 3 — verify and finish.**

````
**thought** Confirm the total landed; re-read F2 from the fresh snapshot.
```cmd
read Sales!F2
```
````

The result shows the computed total. Done.

````
```cmd
done
```
````

---

### Composition variant

Same task using a pipeline to compute the value, then writing the literal:

````
```cmd
let $east = read Sales!A2:C8 | filter region=East
set Sales!F2 = ($east | sum amount)
```
````
