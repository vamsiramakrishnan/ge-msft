# Planning normal form

The canonical shape of a non-trivial program, and where to stop one. This is the structure
[composition-rules.md](composition-rules.md) builds toward; load it when a task is more than a direct
command.

## The four phases

Structure every non-trivial program in four phases, in this order:

```
OBSERVE   read / search / outline            — get the data
DERIVE    let bindings + pure transforms     — compute purely
EFFECT    bounded, concrete host mutations   — the smallest effect set
VERIFY    (optional) read-after-write later  — confirm, in a *later* turn
```

A program that **reads, derives, then emits a tight effect set** is correct by construction; one that
interleaves reads/derives/effects is where bugs live. This is far more useful than "keep it under N
lines" — length is not the signal, phase-order is. `surface_cli normalize` reorders a program into
this form and flags anything that can't be (e.g. a read after an effect).

Note that **VERIFY is a later turn, not the tail of this program** — you cannot read the result of an
effect you emitted in the same program (an effect result is not a value). If the request needs you to
confirm a write, do it after the result block, in the next turn.

## When to break into a new program / turn (semantic boundaries)

Keep one coherent OBSERVE→DERIVE→EFFECT chain in **one** program. Start a **new** program/turn only
when one of these holds — never merely because the script got long:

| Boundary                               | Why it's a break                                                                                                                                                                                |
| -------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Fresh observation required**         | You write, the host recalculates, and you need the _new_ values (write formulas → read computed results → chart). The values don't exist until the first effect lands — a second OBSERVE phase. |
| **Host-minted id consumed downstream** | An effect produces a non-deterministic id you must then target. An effect-result dependency, not pure composition (the deferred Option-B case, ADR-0008 §3).                                    |
| **Approval authority changes**         | An in-document edit and an externally-visible send (mail/post) must not ride one approval just because they share a script.                                                                     |
| **Artifact / surface changes**         | Excel analysis → PowerPoint creation is a signed handoff, not one transaction across two add-in instances.                                                                                      |
| **Failure domains differ**             | A reversible formatting change and an irreversible external post are not one atomic unit.                                                                                                       |
| **Effect budget exceeded**             | A plan that would mint hundreds of comments is partitioned or rejected before approval, even if algebraically valid.                                                                            |

These are exactly the `approvalClass` / `reversible` / dependency boundaries the runtime's effect DAG
encodes — the model expresses the program; the compiler infers the groups; the gate enforces the
boundaries. You don't write `depends-on`; you respect these breaks.
