# ADR-0011: Runtime extension lifecycle and outcome records

**Status:** Implemented in the shared runtime and production task pane, September 2026.

## Context

The repository already had a capable command algebra, an actuation trigger registry, and an
Orchestrator. Production boot did not install those event reactions. Adding a request context
provider, a response validator, or a task outcome check required changes inside `AssistSession`.
An emitted `done` could also hide failed operations, while an observer exception after a successful
write could make that write look failed. These gaps made feature growth expensive and unreliable.

## Decision

Add `RuntimeHooks`, `RuntimeExtension`, and `ExecutionLedger` as shared runtime APIs. Wire lifecycle
hooks through task receipt, model streaming, read/workspace tools, plan review, gated effects, and
task completion. Extensions use namespaced, atomic registration with explicit disposal. Hook payloads
are isolated snapshots; only designated pre-operation and verification phases may block. Request
context providers return validated, budgeted data rather than editing prompt instructions.

The production pane owns the Orchestrator, extension registrations, diagnostics subscriptions, and
page lifecycle cleanup. Host events build context and offer actions; they do not silently run the
assistant. A separate Outlook send runtime installs the same extension definitions independently.

Required checks fail closed on timeout or exception. Post-operation observers fail independently so
successful host receipts remain truthful. Each task has a bounded in-memory metadata record and a
final verification phase. The shipped verifier rejects incomplete work; domain-specific readback
checks can be installed without rewriting the executor. Preserve host inverse receipts for future
explicit undo work, without claiming automatic rollback.

## Consequences

This creates a stable implementation seam for source enrichment, response checks, event-based task
chips, execution diagnostics, and outcome verification. It preserves delegated identity, engine
policy screening, and human approval. Trusted extension code gains no additional permissions.

It does not implement ADR-0010's compute engine, a durable workflow scheduler, a background Graph
notification service, cross-surface orchestration, or semantic verification of every answer. Streams
remain incremental: an end-of-response check cannot retract text already shown. Office writes that
have begun may outlive cancellation, and completed receipts must be retained accurately.

The full phase contract, examples, operation coverage, timing budgets, and known boundaries are in
[RUNTIME-EXTENSIONS.md](RUNTIME-EXTENSIONS.md). Tests use explicitly fake host/network adapters;
live tenant validation remains a release requirement.
