# ADR-0008 — Surface Commander as an algebra + `surface-cli` compiler policy

**Status:** Accepted (2026-06-24) · refines ADR-0004 (command protocol), ADR-0005 (composable
algebra), ADR-0006 (capability closure); **corrects** ADR-0007 §3 (spill semantics). Scope: the
`skill/m365-surface-commander` bundle, the language-manifest emitter, and the runtime plan compiler.

## Context

`m365-surface-commander` is the Gemini Enterprise skill that turns the model's intent into the `cmd`
program the runtime executes. As the host-native write surface grew (ADR-0007 + the cross-surface
catalog, `docs/CAPABILITY-CATALOG.md`, ~85 kinds), two pressures surfaced:

1. **Catalogue drift.** The exact commands available change by host, requirement set, release profile,
   and tenant policy. A skill that is primarily a *static command catalogue* is wrong the moment the
   injected per-turn capability signature differs from it. The Python checker
   (`scripts/parse_commands.py`) is currently **hand-mirrored** from the TS grammar — six potential
   sources of grammar drift (TS contracts, TS runtime, Python parser, grammar docs, capability map,
   golden corpus).
2. **Two latent languages.** ADR-0007 §3 illustrates `let $r = spill … ; table ($r)` (an effect that
   *returns a bindable reference*), but the implemented parser treats `spill` as a **command/effect**
   that terminates computation. Those are different languages and the ambiguity must be resolved.

## Decision

**Make Surface Commander a compact algebraic programming guide + compiler policy — not a second
planner and not a second execution runtime.** The pipeline of authority is:

```
intent → planner (is composition needed?) → Surface Commander (produce cmd source)
       → surface-cli (check / normalize / budget / plan — PURE)
       → runtime parser + compiler (authoritative) → dependency-aware effect plan
       → preview + approval + gate → bridge execution
```

> The skill **produces** source. `surface-cli` **validates** source. The runtime **alone** creates and
> executes effects. The bridge **alone** mutates the host.

### 1 — Surface Commander teaches the algebra, not the catalogue

`SKILL.md` (the only default-loaded file) carries the protocol, the algebraic laws, the decision
procedure, and the output contract. Everything host-specific or pattern-specific is a
progressively-loaded reference. The skill teaches four things:

- **The value algebra** — `Table · Number · Text · Boolean · Selector · RangeRef` (later:
  `DocumentRef · SlideRef · MessageRef · ArtifactRef`).
- **Operator signatures** — pure operators return values and compose freely; **effect** operators
  consume values and **terminate** computation (`set/spill/table/chart/cf`, `suggest/comment/reply`,
  …). The exact signatures are **injected per turn** from the runtime capability registry — `SKILL.md`
  defers to the injected signature, never overrides it.
- **The algebraic laws** — type laws (`filter(p, select(C,T))` valid iff `columns(p) ⊆ C`;
  `head(n, sort(k,T)) ≠ sort(k, head(n,T))`) and operational laws (derive before mutating; reuse a
  binding over a repeated read; emit the smallest effect set; never manufacture a selector/id/range;
  never treat an effect result as a pure value; no hidden cells as scratch memory; never expose a
  capability absent from the injected signature).
- **A composition decision procedure** — inspect the turn's signature → direct command if sufficient →
  read the smallest source → bind reused values → transform purely → compute concrete effect targets →
  emit the minimum dependent effect set → run `surface-cli` past the complexity threshold → output
  exactly one `cmd` program.

### 2 — One skill, progressively disclosed (not six surface skills)

Retain `m365-command-planner` (decides *what*) and **one** `m365-surface-commander` (decides *how to
express it*). Splitting into Excel/Word/… Commanders would create skill-routing failure, duplicated
safety instructions, divergent CLI semantics, and more grammar drift. Structure:

```
m365-surface-commander/
  SKILL.md                    # protocol + laws + decision procedure + output contract (DEFAULT-LOADED)
  references/                 # algebra, composition-rules, planning-normal-form, errors-and-recovery,
                              #   <surface>-semantics.md  (loaded on demand)
  patterns/                   # top-n-report, anomaly-review, evidence-backed-redline, … (reasoning templates)
  scripts/surface_cli.py + surface_cli/{parser,checker,types,normalizer,budget,generated_language}.py
  assets/example-sessions/
```

### 3 — Spill is terminal (Option A). Reserve effect-result refs (Option B) for host-minted ids only

The language invariant is **expressions return values; commands produce effects.** `spill` is a
command; at dry-run the compiler infers the written range from `rows × columns` of the source table —
no runtime bind of the effect result. **ADR-0007 §3's `let $r = spill …` illustration is superseded
by this ADR** and must be corrected to the terminal form.

Option B (`spill : Range × Table → Effect<RangeRef>`, monadic `bind`) is **deferred and scoped**: it
returns only for Plane-B kinds whose target is a *host-generated id that does not exist until the first
effect lands* — `update-message`, `set-reaction`, `graph-patch-page`. It must never be introduced for
the in-document range case, where it buys nothing and complicates dry-run/retry/partial-failure.

### 4 — `surface-cli`: a deterministic compiler tool, never an executor

Add a `surface-cli` helper. It answers: does this parse? does it use only this turn's capabilities?
what are the inferred input/output types? how many effects, and which depend on which? what
ranges/artifacts are touched? does it exceed a policy budget? can it normalize to canonical form?

It **must never**: call Office.js or Graph, acquire tokens, discover capabilities independently,
execute generated code, mutate documents, approve effects, or repair a program semantically without
model review. Subcommands: `check · normalize · explain · budget · plan · simulate` — where
`simulate` is a **pure range/type projection only** (what *would* be touched), never a host-state
simulation.

**One parser authority (anti-drift).** `@ge/contracts` is the single source of truth and emits a
**versioned language manifest** (`dist/language/m365-cli-<v>.json`: verbs, operator signatures, value
types, write-kinds, limits). The skill sandbox runs **Python/Bash**, so `surface_cli.py` is a bounded
**advisory preflight** whose tables are **generated** from that manifest into
`generated_language.py` — never hand-maintained. The **TS runtime parser decides authoritatively**
downstream; the Python preflight only catches structural errors early. The existing
`parse_commands.py` + parity corpus become the **conformance gate**: the corpus proves
`generated-python ≡ TS-grammar` on every release.

### 5 — Break generated programs at semantic boundaries, not by line count

Compile a linear `cmd` program into **one dependency-aware plan**; do not decompose a coherent
read→transform→spill→table/chart family into separate model turns (that causes read drift, approval
fatigue, partial execution). Break into a *new phase* only when:

- a **fresh observation** is required (write → host recalculates → read calculated values);
- an op yields a **nondeterministic host id** consumed downstream (effect-result dependency — the
  Option-B case);
- **approval authority changes** (modify workbook → send external email);
- the **artifact/surface changes** (Excel analysis → PowerPoint creation = a signed handoff);
- **failure domains differ materially** (reversible formatting vs an externally-visible Teams post);
- an **effect budget** is exceeded (a plan that would mint 500 comments is partitioned/rejected
  pre-approval).

The canonical program normal form the skill teaches:
`OBSERVE (read/search) → DERIVE (let + pure transforms) → EFFECT (bounded concrete mutations) →
VERIFY (optional read-after-write in a later phase)`. This replaces any "keep scripts under N lines."

### 6 — The compiler produces a dependency DAG, inferred not authored

The runtime compiles to typed `PlanNode`s (`pure` / `effect`) carrying `dependsOn`, and for effects:
`reads`/`writes` resource sets, `approvalClass`, `reversible`, `idempotencyKey`,
`failurePolicy: stop-dependents | continue-independent`. Dependencies are **inferred** by the compiler
(variable use, overlapping read/write resources, derived ranges, object references, ordering rules) —
the model never writes `depends-on` syntax. On `e1` failure, dependents skip
(`prerequisite_failed`); independent effects may still run if policy permits.

This is a **dependency-aware effect plan / saga with bounded compensation**, **not** an atomic
transaction — and it must not be described as atomic unless every bridge provides a reliable inverse
(the `InverseDescriptor` model is best-effort and includes `not-reversible` kinds).

### 7 — Patterns are reasoning templates; rename in-language `def` to *recipes*

`patterns/` holds algebraic templates (preconditions · pure core · effect core · failure rule ·
anti-patterns) — guidance for reasoning, **not** new hidden high-level commands (that would create a
second opaque DSL over the CLI). Terminology: reserve **skill** for Agent Skill bundles; rename
reusable in-language `def` constructs to **recipes** in all product/user-facing docs.

### 8 — When the model invokes the helper

Direct single effect → generate directly. One pure pipeline + one effect → directly unless syntax is
unfamiliar. >1 binding → `check`. >2 effects → `check + budget`. Any dependent materialization →
`check + plan`. Any recipe `def` → `check + normalize`. Any parser-correction turn → `check`. Near a
policy limit → `check + budget`. Cross-artifact continuation → `plan`. (Cheap actions stay fast;
deterministic machinery runs where models make *structural* errors.)

## Consequences

- **Evolution sequence.** Phase 1 — slim `SKILL.md` to protocol/laws/decision-procedure, demote the
  static verb catalogue, make injected signatures authoritative, lock terminal-effect semantics, rename
  `def`→recipe. Phase 2 — emit the language manifest from `@ge/contracts`; build `surface-cli`
  (`check/normalize/explain/budget/plan`) with generated Python tables. Phase 3 — dependency-aware
  effect plans in the runtime (resource sets, inferred deps, idempotency keys, failure propagation,
  grouped approval preview). Phase 4 — the pattern library + the evaluation matrix.
- **Evaluation (the acceptance bar).** Measure first-pass parse success, capability-violation rate,
  average repair turns, unsupported-verb rate, effect-budget rejection rate, approval-preview
  correction rate, program length, helper latency, and the share of helper invocations that find a
  *real* defect — across three conditions: base model · Surface Commander only · Surface Commander +
  `surface-cli`. The helper ships only if it improves execution quality net of its latency.
- **Budget is defense-in-depth.** `surface-cli budget` advises; the runtime gate re-checks. The skill
  is never the only thing between "algebraically valid" and "500 comments."
- **Corrects ADR-0007 §3.** The effect-result spill illustration there is non-normative; this ADR's
  terminal-effect form is authoritative.

## The durable abstraction

Surface Commander teaches the algebra · `surface-cli` proves the program is coherent · the runtime
proves it is permitted · the bridge alone performs the mutation.
