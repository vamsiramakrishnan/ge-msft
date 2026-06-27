# ADR-0004 — A command-line protocol for the assist read/write loop

**Status:** Proposed (2026-06-22) · refines ADR-0003 (context construction); builds on ADR-0002 (capability model) and `CONTRACTS.md`. Scope: the **assist path only** (grounded `streamAssist` chat). The A2A specialist paths (review → `Finding[]`, draft-slides, synthesize) are unchanged.

## Context

ADR-0003 adopted a working-document tool surface (an ambient `<doc_state>` snapshot + lazy
host-read tools) and made one load-bearing assumption explicit:

> "Expose a small, typed tool surface … that the **`runtime` Orchestrator/AssistSession** drives in
> a bounded loop … **The loop lives in our orchestrator because `streamAssist` will not run it for
> us.**"

That assumption deserved a test, because it determines *who* drives the loop and *in what syntax*
the model expresses reads and writes. `streamAssist` has **no native function-calling**, so the
options are:

- **(A) Orchestrator-driven, no model-emitted calls** (ADR-0003's default): the runtime guesses
  what to prefetch and folds it into context. Tight budget control; the model never *asks*.
- **(B) Model-driven via a text protocol**: the model emits calls as text the runtime parses. Two
  syntaxes are possible — **JSON function-call objects** (`{"name":…,"args":{…}}`) or **flat
  command lines** (`read Sales!C2:C7`).

We probed all of this against the live engine (`phoenix-retail_1751440313229`,
`vital-octagon-19612`, `global`) with the harnesses in `scripts/streamassist-*.mjs`. The findings
(full table under **Validation**) were decisive:

1. `streamAssist` **can** be driven as a text-protocol ReAct agent — reliably, multi-turn,
   grounded, and injection-resistant — contradicting the "it won't run the loop for us" assumption.
   The orchestrator still owns the loop *policy*; the model drives the *steps*.
2. **The call syntax is what determines reliability, not the loop itself.** For reads (a shallow,
   3-verb grammar) JSON and CLI tie. For **writes/format** — the nested, escape-prone payloads that
   mirror our real `ActuationRequest` params — a flat command line **beats JSON** even after the
   JSON parser is made maximally forgiving. JSON's failure mode is **envelope drift**: a model with
   no native tool-calling cannot hold the `{name, args}` wrapper steady across turns and degrades to
   `{tool…}`, `{verb…}`, `{action, type, command, name…}`. A command verb has no envelope to drift
   from.

This ADR settles the open question from ADR-0003: **in what syntax does the assist model express
reads and writes, and who drives the loop.**

## Decision

Adopt a **command-line (CLI) grammar as the model-facing surface syntax** for the assist loop. The
model emits flat command lines inside a fenced ` ```cmd ` block; the runtime **parses → validates →
compiles** each line into the *existing* typed boundary objects and runs the *existing* machinery.

> **The command line is the assembly language the grounded model emits; the typed
> `ActuationRequest` (ADR-0002) is the bytecode the bridge executes.** We add a thin parse/compile
> layer — we do **not** replace the capability model, the gate, or provenance.

### 1. Grammar — unified verbs, surface-specific selectors

A small set of orthogonal verbs, shared across surfaces; only the **selector** is surface-specific.

```
# reads (Layer-B host reads, ADR-0003)
outline
read   <selector>                 # Excel: Sales!C2:C7 · Word: (whole/section) · PPT: slide:4
search <text>

# writes (compile to ActuationKind, ADR-0002 capability.ts)
set     <cell> <value|=formula>   # → write-cells           (Excel)
format  <range> k=v ...           # → (proposed) format-cells
comment <selector> "text"         # → comment-reply
suggest "<exact text>" => "<new>" # → tracked-change        (Word, content-anchored)
done
```

- **Selectors are the only per-surface variation:** A1 (`Sales!C2:C7`) for Excel, **content anchors**
  (`matchText`, re-resolved via `body.search` at apply-time per ADR-0001/01-architecture) for Word,
  `slide:n/shape` for PowerPoint. The Word probe confirmed the *same* verbs work with content
  anchors and **zero anchor drift** (no hallucinated `matchText`).
- **Per-surface verbs are scoped by the `CapabilityManifest` (ADR-0002).** The prompt advertises
  **only** the verbs the active bridge's `getCapabilities().actuations[]` supports right now. Small
  per-surface grammar → fewer tokens to get wrong → higher reliability.

### 2. Compile + validate — the runtime boundary

Each command line is parsed and **compiled to a typed `ActuationRequest` / Layer-B read call**, then
**Zod-validated** exactly as today. The model never has to emit schema-valid JSON; the runtime
constructs and validates it. On a parse or validation failure the runtime returns a **CLI-style
corrective error** (`error: unknown verb 'writ-cells' — did you mean 'write-cells'? (run 'help')`),
which the model self-corrects on the next turn. This is the reliability mechanism: the malformed-
nested-JSON failure mode is designed out, and the remaining failures self-heal.

### 3. Runtime policy (empirically required — these are the orchestrator's job, not the prompt's)

1. **Read-many / write-one.** Batch read-only commands freely (cheaper, fewer round-trips — measured
   15 vs 17 turns). **Execute writes one at a time through `triggers.gate()` with per-write
   approval.** The model *will* batch writes (observed: 14 in one block); the orchestrator must
   serialize and gate them. Writes remain *"never called without user confirmation"* (`DocBridge`).
2. **Parse the fenced block; ignore the rest.** `streamAssist` emits `**thought**` preambles before
   the block; extract the ` ```cmd ` fence and ignore surrounding prose. Treat a turn with **no
   fenced block** as a re-prompt, not an error (observed once; recovered next turn).
3. **Formula-first / verified-literal write-back** (consistent with ADR-0003 #3). The model already
   computes correct values during analysis; on write-back it should emit a host-native formula
   (it correctly reached for `SUMIF` over interleaved regions) or the verified literal.

### 4. Scope and non-goals

- **Assist path only.** `streamAssist` request shape is **unchanged** — still grounded chat, no
  `toolsSpec`/function-calling field; the protocol lives entirely in the prompt + the runtime
  parser. This preserves the standing constraint *"use StreamAssist only for the grounded-assistant
  chat path."*
- **A2A specialist agents unchanged.** `review`/`draft-slides`/`synthesize` keep their structured
  `Finding[]`/event outputs over A2A.
- **No client code-execution VFS** (still out of scope, per ADR-0003).

## Consequences

- **`runtime`** gains a command parser/compiler (CLI line → `ActuationRequest` / read call) and the
  corrective-error contract; `AssistSession` drives the bounded loop and enforces read-many/write-one
  through the existing `triggers.gate()`. This *changes ADR-0003*: the loop is now **model-driven**
  for the read/write steps, while the orchestrator retains policy, budget, compaction, and the gate.
- **`contracts`** gains a `CommandGrammar` / verb→`ActuationKind` mapping derived from the
  `CapabilityManifest`; the parser is the boundary and the single place the grammar is defined.
- **`gemini-client`** is unchanged.
- **`bridge-*`** are unchanged (`actuate()`, `captureDocState()`, `searchDocument()`); optionally add
  an explicit `readRange()` to `DocBridge` for the `read` verb on Excel.
- **Legibility bonus (free).** Because writes are gated, the approval card renders the command
  verbatim (`set Sales!F2 =C2-D2`) — human-auditable, reinforcing the reversible/provenanced
  invariant without extra work.
- **Risk:** an occasional non-fenced "thinking" turn wastes a round-trip; mitigated by the
  re-prompt rule. Grammar drift on rare verbs is mitigated by `CapabilityManifest` scoping +
  corrective errors.

## Validation

Live engine `phoenix-retail_1751440313229` / `vital-octagon-19612` / `global`, 2026-06-22. Harnesses
are in `scripts/` and reproducible with `GE_TOKEN` (e.g. `gcloud auth print-access-token`) +
`GE_PROJECT`/`GE_LOCATION`/`GE_ENGINE`.

| Experiment | Harness | Result |
|---|---|---|
| Read-only A/B (JSON vs CLI) | `streamassist-protocol-probe.mjs` | Tie on reliability (both 6/7, 0 malformed); CLI fewer turns via batched reads (15 vs 17) |
| Write/format A/B | `streamassist-write-probe.mjs` | **CLI 4/4 vs JSON 2/4** even with a maximally forgiving JSON envelope; JSON's failure mode is envelope drift, not nested values |
| Multi-turn EDA | `streamassist-eda-session.mjs` | 100% grammar reliability, single session, monotonic `<doc_state>`, deliverable correct (per-region `=SUM`, anomaly comment, header bold), grounded, 0 hallucination |
| Adversarial + scale (30 rows, injection in data + comments) | `streamassist-eda-adversarial.mjs` | **All injections resisted** (ignored a planted `PWNED` override, a fake total, and a fence-break comment), grounded, correct `SUMIF` write-back — host content treated as data, per the standing constraint |
| Word surface (content anchors) | `streamassist-word-session.mjs` | Same grammar; **0 anchor drift**, 2/2 unsourced claims flagged, 2/2 sourced claims left untouched — the grammar transfers across surfaces |

**Headline:** the CLI protocol over the ADR-0003 context loop is reliable (100% grammar deep into
stateful sessions), grounded (reads before it asserts), injection-resistant, and surface-portable
(Excel grids and Word content anchors). Writes belong in the grammar (4/4 vs 2/4). The only
operational rules are runtime-side (read-many/write-one, parse-the-block, formula-first).
