# ADR-0005 — A composable capability algebra for the assist loop

**Status:** Accepted (2026-06-22); Phases 1–3 implemented (2026-06-23) · refines ADR-0004 (command protocol); builds on ADR-0003 (context construction), ADR-0002 (capability model). Scope: the assist path (grounded `streamAssist`). A2A specialist agents unchanged.

## Context

ADR-0004 gave the assist loop a flat command grammar — the model emits one verb per turn and
carries the orchestration itself (read → reason → write → read…). That is a REPL. It works, but:

- **Complex tasks cost many turns.** A per-row formula fill or a multi-step reconciliation is N
  gated turns; the model rebuilds intent each turn.
- **Nothing is reusable.** Every task is re-derived from primitives; there is no way to name and
  replay "the quarterly reconcile" or "redline a contract".
- **Reads can't feed writes without a round-trip.** The model reads, the result comes back as text,
  it reasons, it writes — the data never flows directly from a read into an actuation.

We want capabilities to **compose**: the model should be able to *program* the document — pipe a
read through pure transforms, bind intermediate values, and apply a batch of effects — and we should
be able to name a composition and reuse it. This turns `streamAssist` from a thing that *answers*
into a thing that *actuates, controls, and reads* as a system, alongside the add-in's own directed
(UI-initiated) actions.

## Decision

Introduce a **composable capability algebra** over the ADR-0004 grammar, anchored on one keystone:

> **A typed *value* layer between reads and actuations.** Reads produce **values** (`Table`,
> `Number`, `Text`). Pure transforms compose values (`filter`, `select`, `sum`, `sort`, …).
> Actuations *consume* values and produce **gated `Effect`s**. **Pure composition runs freely; only
> the typed `Effect` terminals actuate** — each through the existing gate (ADR-0004) +
> `isUnsafeFormula` + provenance.

That single split — *pure composes freely, effects gate* — is both the composition mechanism and the
safety boundary. It is ADR-0003's "code-execution is for I/O, not analysis" expressed as a type
system.

### 1. The capability registry (unification)
A `Capability` is `{ name, signature, kind: read | pure | effect, compile, gatePolicy, provenance }`,
derived from the surface's `CapabilityManifest` (ADR-0002). The add-in's **ribbon buttons (directed)**,
the **model's commands (autonomous)**, and **saved skills (programmatic)** all become front-ends to
this one registry. There is one algebra; three ways to invoke it.

### 2. The composition surface (a deliberately constrained shell)
On top of ADR-0004's lines: **pipes** (`read X | filter … | sum …`), **bindings** (`let $m = …`),
**bounded iteration** (`for row in rows($t): …`), and **named definitions** (`def reconcile(a,b): …`)
— the last being a reusable capability/skill. It is **not** a general programming language: loops are
bounded, there is no arbitrary code-execution, and effects exist only as gated actuations. The
constraint is the product — legibility and the gate are what make autonomous actuation trustworthy.

### 3. Planner / executor
The model emits a **plan** (a composition). The runtime: **type-checks** it against the
`CapabilityManifest` (unsupported verb / non-composing types fail before any effect) → **dry-runs**
(executes the reads + pure transforms, computes the effect-set without writing) → **previews** the
effect-set for one **plan-level approval** (extending ADR-0004's fail-closed per-write approval) →
**executes** (reads batched, effect terminals gated one-by-one, provenance persisted). The plan is a
first-class artifact: inspectable, dry-runnable, replayable, and saveable as a skill.

A declarative plan emitted once and type-checked **drifts even less** than ADR-0004's per-turn
commands (no per-turn envelope to reconstruct) — the same reliability argument, one level up.

### 4. Skills = saved compositions (the compounding library)
A skill is a named, parameterized composition, stored per-user/org (the bootstrap-config pattern from
the Claude-for-M365 teardown — but ours are **typed, gated, grounded document programs**, not prompt
config). The model can call a skill or define one; the org's capability set grows **without shipping
code**.

## Safety invariants
- **Pure/effect type split** — a malicious composition (model output is shaped by untrusted document
  content) can read and compute but cannot silently write; every write is an `Effect` terminal that
  hits the gate + formula guard + provenance.
- **Plan-level gate + dry-run** — approve the program and its full effect-set, with a preview of
  exactly what will change, before anything changes.
- **Type-check against `CapabilityManifest`** — no half-applied plans.
- **Bounded execution** — bounded loops, a per-plan effect cap (extends ADR-0004's per-turn caps).
- The ADR-0004 untrusted-input boundary is **preserved**: the planner parses/type-checks/gates the
  same way — composition moves the gate to *plan* granularity, it does not weaken it.

## Phased path
1. **Pure value layer + pipes/bindings over reads** *(implemented)* — `read | filter | sum` → a
   value; bind with `let`. Almost entirely pure ⇒ low risk, immediate analytical power, no new
   effect surface. (`packages/contracts/src/expr-grammar.ts` parser; `packages/runtime/src/compose.ts`
   value model + transform registry + evaluator; `AssistSession.runCommands` evaluates composed
   read-expressions in the loop.)
2. **Plan = composed effect-set + dry-run + one approval** *(implemented)* — a turn's effects form a
   PLAN: type-check (verb's kind ∈ `manifest.actuations`, `$vars` bound, effect-arg expressions
   parse) → dry-run (execute reads + pure, then RESOLVE each effect — evaluate any expression arg to
   a `Value`, render it to the param, `compileCommand` → a Zod-validated `ActuationRequest` — **with
   zero actuation**) → preview the effect-set (`plan-preview` event) → **one** plan-level approval
   (`approvePlan`, fail-closed: no approver ⇒ the whole plan is blocked) → gated execution (each
   effect through the existing gate + `isUnsafeFormula` + provenance; per-plan effect cap retained;
   `changeId` minted once at dry-run). The **keystone connection**: an effect's value/text arg may
   *consume* a composed value — `set Summary!B2 = ($anz | sum Revenue)`, `set B3 = $total`. The
   expression layer stays pure — a `$var | set …` inside an arg is rejected, never executed. Plain
   `ask()` and the ADR-0004 standalone-effect path (per-write `approveWrite`, Track A) are unchanged
   and remain the fallback when `approvePlan` is absent.
   (`packages/contracts/src/{expr-grammar,command-grammar}.ts`;
   `packages/runtime/src/assist-session.ts` `resolveEffect`/`executePlan`/`PlanEffect`.)
3. **Named skills** *(implemented — parameterized macros)* — define (`def name(p…): … end`) / call
   (`name arg…`) parameterized compositions; the compounding library. A `def` registers an in-session
   skill (no execution → a confirmation); a call binds args→params, substitutes `$param` tokens, and
   runs the expanded lines through the **same Phase-2 plan** (type-check → dry-run → `approvePlan` →
   gated execute) — a skill call is just a plan, with no new gate/approval bypass. Substitution is
   textual into already-parsed entries: an arg cannot inject a command line (newline/fence rejected),
   only whole declared `$param` tokens substitute, and the expansion is bounded (command budget, write
   cap, body cap, nesting depth). A name may not shadow a built-in verb. `for`/`each` iteration and
   durable (host-metadata) skill persistence are deferred to a later phase.
   (`packages/contracts/src/skill-grammar.ts` grammar/AST; `packages/runtime/src/skill-registry.ts`
   registry + expansion; `AssistSession.processEntry` registers defs / expands calls into the plan.)
4. **Cross-surface compositions** — read Excel → write PowerPoint; the unit, programmable end-to-end.

## Consequences
- **`contracts`** — the command grammar gains an expression layer (pipelines + bindings → a typed
  `ParsedExpr` AST); `Value` schemas. The grammar stays the single source of truth.
- **`runtime`** — a `Value` model + a pure transform registry + an **evaluator** (dispatches `read`
  to the `DocBridge`, applies pure transforms, holds the binding env); `AssistSession` evaluates
  composed read-expressions in the loop and (Phase 2+) plans/executes composed effects through the
  existing gate. Plain `ask()` and the ADR-0004 simple-command path are unchanged.
- **`bridge-*`** — unchanged for Phase 1 (reads are parsed into `Table` values in the evaluator);
  later phases may add a structured read for fidelity.
- **`gemini-client`** — unchanged.

## What stays
The typed `ActuationRequest` boundary, the actuation gate + fail-closed approval, durable provenance,
the untrusted-content boundary, content-anchored Word writes, residency, and identity federation —
all unchanged. ADR-0005 is additive: ADR-0003 made the document an environment, ADR-0004 gave it a
command language, ADR-0005 makes that language compose.
