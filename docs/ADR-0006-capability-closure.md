# ADR-0006 — Capability closure: one descriptor, derived surfaces, conformance-enforced

**Status:** Accepted (2026-06-23) · completes ADR-0002 (capability model); precondition for ADR-0005 Phase 2 (plan type-check/dry-run is only sound if the manifest is truthful).

## Context

An external review (and our own `CAPABILITY-MAP.md`) found that the *executable* capability set —

```
declared manifest ∩ Office APIs on this host ∩ CLI compiler ∩ bridge actuate() ∩ user policy
```

— is **never computed**, and the independent representations have **drifted** (verified in code, 2026-06-23):

- **Word** advertises `insert-ooxml` and `fill-content-control`, but `actuate()` handles neither — **phantom capabilities**.
- **Outlook** advertises `create-mail`, but the bridge handles only `reply-mail` — phantom.
- **Excel** advertises the `read <A1>` verb path, but `ExcelBridge` has **no `readRange()`** — so an ADR-0005 Phase-1 pipeline that starts from a cell read **fails on Excel today**.
- Several bridge effects (`insert-slide`, `append-page`, `comment-reply`, `post-message`, …) are handled but have **no CLI verb** — unreachable from the model.
- Every `getCapabilities()` returns a **static** manifest despite comments promising runtime detection.

A model that *composes* capabilities (ADR-0005) over a non-closed set composes **phantoms**, and a
plan-level approval over phantom capabilities is not trustworthy. Closure is therefore a
prerequisite, not hygiene.

## Decision

**Make the capability set closed and truthful, and enforce it with conformance tests that fail the
build on drift.** Two severities:

- **Phantom (a lie) → hard build failure.** An advertised actuation kind that `actuate()` does not
  handle, or an advertised read verb with no bridge read port. Resolve by **implementing it** (when
  cheap) or **un-advertising it** (preferred when not). A surface must never claim what it cannot do.
- **Gap (unreached) → tracked, not fatal.** A bridge-handled kind with no CLI verb. Recorded against
  an explicit allow-list so it is visible and burned down deliberately, without blocking on full
  parity in one wave.

Mechanism:

1. **A pure `checkCapabilityClosure({ manifest, handledKinds, readPorts, verbKinds })` helper** in
   `@ge/contracts` returning `{ phantoms, unreachedReads, gaps }` — the single definition of closure.
2. **Per-surface conformance tests** assert `phantoms == [] && unreachedReads == []` (hard), and
   compare `gaps` to a checked-in allow-list (tracked).
3. **A `Capability` descriptor** (`{ name, surface, kind: read|pure|effect, signature, compile,
   gatePolicy }`) is introduced as the forward source of truth from which the manifest, the
   verb→kind map, and dispatch are derived for *new* capabilities; existing ones are migrated
   opportunistically. (Full registry migration is incremental — the conformance gate is what makes
   the incremental path safe.)

This wave also fixes the verified drift (un-advertise/implement the phantoms; add Excel `readRange`;
expose `comment-reply` as the `reply` verb) so the conformance gate goes green.

## Scope (this wave)
- Closure helper + conformance tests + the `Capability` descriptor scaffold (contracts/runtime).
- Drift remediation in the bridges; Excel `readRange`; the `reply → comment-reply` verb.
- Broader CLI parity (`slide`/`page`/`mail`/`post` verbs, other surface read ports) is **tracked as
  gaps**, not delivered here.
- Runtime requirement-set *detection* in `getCapabilities()` stays a follow-up; manifests remain
  static but are now conformance-checked.

## Consequences
- `contracts` — the closure helper + `Capability` descriptor; the verb↔kind map stays the grammar's
  source of truth and is one input to closure.
- `runtime` — the `reply` verb compiles to `comment-reply`; `read` dispatches to the new Excel port.
- `bridge-*` — un-advertise/implement phantoms; add `readRange`; per-bridge conformance tests.
- Conformance failures now block merge — drift cannot silently return.

## What stays
The CLI-as-source-language / `ActuationRequest`-as-IR separation (ADR-0004), the gate + fail-closed
approval, durable provenance, the untrusted-content boundary, ADR-0005 Phase 1 composition — all
unchanged. ADR-0006 makes the capabilities they operate over honest.
