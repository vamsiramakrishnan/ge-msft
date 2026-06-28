---
title: Errors And Recovery
kind: reference
skill: m365-surface-commander
topics: [corrections, parser-errors, stale-anchors, budget]
load_when: A result block contains an error or a command failed validation.
---

# Errors and recovery

When a command is wrong you get a short, CLI-style correction in the ` ```result ` block — not a
failure. Read it, fix that line, and continue. The whole loop is designed to be self-correcting; you
never need to apologize or restart, just repair.

## The corrective contract

Each malformed line yields one `error: …` entry, in order, aligned to the command that produced it.
A valid line still runs; one bad line does not poison the others — **except** that a block carrying
any parse error is never honored as `done` (fail-closed: you must clear the error first).

## The error families and how to recover

| You see                                                             | What it means                                             | Recover by                                                                                                     |
| ------------------------------------------------------------------- | --------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| `unknown verb "writ" — did you mean "write"?`                       | a typo'd or unavailable verb                              | use the suggestion; if it's an _unavailable_ verb, it isn't in this turn's `<capabilities>` — pick one that is |
| `unknown capability "/insert-imag" — did you mean "/insert-image"?` | a `/`-surface kind name typo                              | use the suggested `ActuationKind`                                                                              |
| `… is not supported on this surface`                                | the verb/`/<kind>` isn't advertised this turn             | the surface can't do it — choose an advertised capability or stop (`done`) and say why                         |
| `set needs a cell and a value …`                                    | missing/garbled arguments                                 | re-emit with the usage shown                                                                                   |
| `spill needs a composed table, not a literal`                       | you passed a literal where a table expression is required | bind a table and spill it (`spill R = ($t)`), or use `set` for one scalar                                      |
| `$x used before it is bound`                                        | a `$var` referenced before its `let`                      | move the `let` above its use (`normalize` does this)                                                           |
| a degraded annotation (anchor drifted)                              | the exact text you anchored on changed                    | re-`read`/`search` for the current text and re-anchor; never anchor on guessed text                            |

## Stale anchors degrade, they don't break

Content-anchored writes (`suggest`, `comment`, Word `/`-kinds) re-resolve their anchor at apply-time.
If the exact text has drifted, the change **degrades to a panel item** rather than landing a broken
annotation — prefer that to forcing a write on approximate text. If you see a degrade, re-read and
re-anchor on the text as it now reads.

## Budget rejections

If a program would exceed the effect/read/cell budget, it is rejected **before** approval (not
half-applied). Partition the work (e.g. one `cf` highlight instead of 500 comments) or split across
turns at a semantic boundary. Run `surface_cli budget` first when you're near a limit.

## Preflight to avoid the round-trip

Most of the above is catchable _before_ you emit, with `surface_cli check` (unknown/out-of-signature
verbs, unbound `$vars`, wrong dependencies) and `surface_cli budget`. The runtime parser is still the
authority — the preflight just saves a correction turn. See the invocation policy in
[SKILL.md](../SKILL.md).
