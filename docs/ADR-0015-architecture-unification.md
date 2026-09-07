# ADR-0015 — One owner for each execution boundary

**Status:** Accepted for implementation (September 2026). Scope: architecture cleanup and boundary
correctness; no new product features or protocol delimiters.

## Context

The six bridges duplicated switch dispatch and handled-capability lists. Analysis schemas and
provenance records had multiple owners. The controller repeated execution setup/cleanup across
routes, while command observation storage and mode policy were mixed into the main session loop.
Python skill tooling independently maintained vocabulary and validation assumptions; missing
metadata could silently weaken checks. Architecture guidance still included superseded gateway
requirements.

These were correctness problems as well as maintenance costs: duplicated paths disagreed about
capability admission, cancellation, receipt correlation, provenance, and stream completion.

## Decision

| Responsibility | One owner | Deliberately retained boundary |
| --- | --- | --- |
| Payload shapes and pure outcome policy | `@ge/contracts` | Host execution and SQL stay outside contracts |
| Host handler coverage | Executable per-bridge table plus shared dispatcher | Host API quirks, native storage, mutation/readback semantics |
| Runtime request admission and receipt validation | Shared admission, recovery entry, and outcome helpers | Effective session manifest remains authoritative |
| Command disclosure and retained observations | `CommandContextSession` | AssistSession still captures host state and calls the provider |
| Pane execution ownership | Controller acquisition/release helpers | Explicit planner handoffs and route-specific events |
| Skill vocabulary and guard metadata | Generated TypeScript contract manifests | Python provides standalone advisory preflight |
| Skill ZIP contents | Shared source inventory | Each bundle remains independently installable |
| Dependency direction | AST-based repository conformance | Shell remains the adapter composition root |

This is consolidation around the existing interfaces, not a universal adapter framework. The
capability registry remains descriptive; executable handlers and the effective capability manifest
are independently checked. Existing runtime/compute schema exports and Word/Excel record exports
remain as compatibility re-exports. Word XML and Excel JSON storage stay byte-compatible.

## Correctness changes

- Explicit proposals now obey effective capability filters and re-admission before dispatch.
- Receipts are schema-checked and correlated to the actual request before recovery classification.
  Contradictory or uncorrelated success claims become unknown outcomes.
- UI cancellation reaches proposal execution before dispatch. After dispatch, the actual receipt
  survives cancellation; the task can be cancelled while a host effect still landed.
- Empty or malformed cell grids fail validation without refinement exceptions.
- Pending context is committed only after successful provider completion. EOF without `done`
  marks the task incomplete and cannot execute an otherwise complete command fence.
- Proposals use explicit attribution; history preserves verification/recovery information and
  protects stored records from consumer mutation.
- Missing/corrupt CLI metadata fails closed. Unsupported Zod checks fail generation instead of
  silently disappearing from the standalone checker.
- Planner examples and Python validation use the runtime's supported vocabulary. Previous
  Python-only workflow directives are removed; cross-surface guidance remains ordinary plan text.
- Outlook partial/unconfirmed draft changes and PowerPoint failures after queuing a mutation
  retain uncertainty instead of claiming a safe pre-write failure.

## Conformance and maintenance

The repository test checks parsed imports, direct dependency declarations, dependency direction,
and TypeScript reference agreement. Bridge tests compare advertised capabilities with actual table
keys. Shared schema identity and provenance encoding tests protect compatibility. Python parity,
manifest failure cases, deterministic archive/source parity, and generated-resource checks prevent
skills from developing a separate language.

Validation also covers cancellation before and after host dispatch, late controller events,
reentrant state subscribers, pending approvals, provider failure/EOF, and malformed receipts.
Mocked host/provider tests prove local boundary behavior; live Office and tenant acceptance remain
separate checks. No measured latency or universal reliability multiplier is claimed for this refactor.

## Consequences

Adding a host effect means implementing one handler and updating the canonical contracts/discovery
mapping deliberately. Changing an analysis shape or approval classification updates both runtime
and generated preflight metadata. Adding a controller route reuses execution ownership rather than
copying busy/cancel/finally logic. New package edges fail conformance until their intended direction
and declarations agree.

The command grammar, existing delimiters, deterministic SDK programs, and product workflows remain
in place. Future feature work can build on these smaller, explicit boundaries.
