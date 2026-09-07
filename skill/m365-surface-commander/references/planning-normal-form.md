---
title: Planning Normal Form
kind: reference
skill: m365-surface-commander
topics: [observe, derive, effect, verify, approval-boundaries]
load_when: A task needs bindings, multiple effects, or a decision about model-turn boundaries.
---

# Planning normal form

A program observes inputs, derives results, prepares concrete effects, then requests execution.
The host handles preview, approval, freshness checks, effect ordering, and supported readback.
Verification is a host phase; a separate model inference is necessary only for an unresolved decision.

1. **Observe:** read/search/inspect the needed host data, or capture versioned table artifacts.
2. **Derive:** bind pure expressions and analysis artifacts; reuse bindings in dependent operations.
3. **Effect:** materialize the smallest concrete effect set into explicit host targets.
4. **Verify:** the host compares supported effects with readback. `finish when=verified` makes
   successful verification terminal, without another inference just to emit `done`.

`surface_cli.py plan`/`normalize` expose ordering constraints; they cannot grant approval or repair
semantic ambiguity. Reads after effects may require another program because preview preparation
must observe the state against which the user will approve a write.

## One-program reconciliation

With known ranges, column mappings, and destination, emit:

```text
let $invoices = analyze {"kind":"capture","range":"Invoices!A1:C500","headers":true}
let $payments = analyze {"kind":"capture","range":"Payments!A1:C700","headers":true}
let $result = analyze {"kind":"reconcile","spec":{"left":"$invoices","right":"$payments","leftKey":0,"rightKey":0,"leftAmount":1,"rightAmount":1,"leftCurrency":2,"rightCurrency":2,"tolerance":"0.01"}}
analyze {"kind":"materialize","id":"$result","destination":"Results!A1"}
finish when=verified
```

Capture returns immutable artifact references; analysis bindings resolve only in fields whose
schemas expect them. The host resolves dependencies, computes the actual values, prepares the
exact destination mutation, then asks for approval. Never substitute guessed artifact IDs.
The examples' ranges and mappings must be replaced with observed live sources and approved task
parameters. If the mapping is unknown, ask/read enough to resolve it before preparing the write.

Ordinary `let` bindings continue to hold pure expression values:

```text
let $rows = read Sales!A1:B9 | filter region=East
set Summary!B2 = ($rows | sum amount)
finish when=verified
```

Pipelines cannot contain host effects. Analysis artifacts are handles, not table-expression values;
use analyze operations to consume them. `help analyze` supplies action schemas and restrictions.

## Completion is a condition

`finish when=verified` does not declare success. It asks the host to finish only if all effects
have verified and no error, skipped effect, or pending recovery remains. Cancellation or a failed
check cannot become successful completion. Hosts/effects without verified readback cannot satisfy
this condition; inspect their outcome and use legacy `done` only when the task's actual completion
criteria are met. Never describe an unsupported verification as completed.

## Break at a real decision or authority boundary

| Boundary | Reason for another program or host phase |
| --- | --- |
| Fresh calculated values after a write | New values do not exist until the effect lands; observe them afterward. |
| Effect-created ID without a supported binding | Read the actual effect receipt before targeting the new object. |
| Ambiguous source, mapping, or destination | Resolve the semantic choice before preparing a mutation. |
| Different approval authority or failure domain | Do not combine a document edit and an externally visible action under incidental batching. |
| Different surface | Use typed handoff; separate add-in instances are not one Office transaction. |
| Effect/resource budget exceeded | Partition or reduce work before execution; never weaken policy limits. |

Do not split solely because the program is long, because a local operation minted an artifact ID,
or because readback occurs after a write. Preserve dependencies in the program; the compiler and
host enforce effect gates.
