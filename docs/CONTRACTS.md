# Contracts

The executable TypeScript types and Zod schemas are authoritative. This document maps their
owners and explains boundary invariants; it does not duplicate their declarations. Change a
contract in its owning module, update its consumers, and regenerate skill metadata together.

The application is client-direct. Historical gateway routes and `services/agents` examples in the
original design are not supported application endpoints. See [ADR-0001](ADR-0001-client-direct-architecture.md)
and [the unification decision](ADR-0015-architecture-unification.md).

## Canonical owners

Paths below are relative to the repository root.

| Boundary | Authoritative source | Consumers |
| --- | --- | --- |
| Intent, scope, grounding | `packages/contracts/src/intent.ts`, `grounding.ts`, `command-palette.ts` | Pane routes, quick actions, planner |
| Research unit and request | `packages/contracts/src/unit.ts`, `request.ts` | Runtime and provider request construction |
| Findings, sources, stream events | `packages/contracts/src/finding.ts`, `sse.ts` | Provider normalization, runtime, pane |
| Context references and snapshots | `packages/contracts/src/context.ts`, `doc-state.ts` | Host capture, context model, navigation |
| Actuation request/result and capability manifest | `packages/contracts/src/capability.ts` | Compiler, runtime admission, bridges |
| Outcome assessment and receipt validation | `packages/contracts/src/actuation-outcome.ts` | Recovery, runtime completion, pane history |
| Provenance payload and record | `packages/contracts/src/provenance.ts`, `provenance-record.ts` | Runtime attribution, Word/Excel persistence |
| Cell data, analysis actions and programs | `packages/contracts/src/analysis.ts`, `analysis-actions.ts`, `analysis-program-schema.ts` | CLI metadata, runtime workspace, compute, SDK |
| Command, expression, macro, planner grammar | `packages/contracts/src/command-grammar.ts`, `expr-grammar.ts`, `skill-grammar.ts`, `command-plan.ts` | TypeScript parser/compiler and generated Python vocabulary |
| Command discovery and preflight metadata | `packages/contracts/src/language-manifest.ts`, `preflight-metadata.ts` | CLI help/checker, packaged skills |
| Approval classes and plan graph | `packages/contracts/src/plan-graph.ts` | Compiler, runtime, generated CLI checks |
| Capability closure | `packages/contracts/src/capability-closure.ts` | Per-surface conformance tests |
| Host adapter interface and dispatch | `packages/runtime/src/bridge.ts`, `bridge-dispatch.ts` | Six host adapters |
| Runtime hooks and task lifecycle | `packages/runtime/src/hooks.ts`, `execution-ledger.ts`, `assist-session.ts` | Compiled extensions and controller |
| Provider transport | `packages/gemini-client/src/config.ts`, `stream-assist.ts` | Runtime and application composition |

Runtime/compute compatibility exports for moved analysis schemas remain available. New consumers
should import schemas and pure policy from `@ge/contracts`; SQL execution stays in `@ge/compute`.

## Request and stream lifecycle

The client federates the signed-in user's Entra identity through WIF and sends the resulting
short-lived Google bearer token to the configured Discovery Engine transport. Chat, planner, and
command routes select their own skill configuration. Provider JSON/SSE frames are normalized into
the internal `SseEvent` union; the internal event schema is not a claim that the public API emits
our application event names verbatim.

A provider stream must complete successfully before the runtime commits pending context delivery.
An error, policy block, cancellation, or EOF without `done` leaves it pending. EOF also marks the
task incomplete and stops command parsing, even when a complete `cmd` fence was already received.

Command/planner calls use the configured session mode. Ordinary chat retains its conversation;
command context supports a bounded state projection, transcript compatibility, and explicit
conversation compatibility. Session-bound uploads must follow the documented mode constraints.
See [command performance and session modes](COMMAND-PERFORMANCE.md).

## One host effect boundary

All runtime host effects, including explicit proposals and typed SDK programs, follow the same
admission and receipt rules:

1. Parse the request and match its surface to the bridge and current effective capability manifest.
2. Prepare supported recovery state and obtain the route's explicit approval. Missing approval blocks execution.
3. Run effect hooks and trigger gates; re-admit against current capabilities and check cancellation immediately before dispatch.
4. Dispatch through the bridge's actual handler table. Host APIs and mutation semantics remain inside that adapter.
5. Validate the receipt's schema, kind, change ID, and consistency before recovery or completion classifies it.
6. Record the actual outcome and explicit provenance. Cancellation cannot erase a write that already landed.

`assessActuationResult` distinguishes verified, unverified, uncertain, rejected, and failed outcomes.
A generic host success without readback is unverified; verified program completion requires actual
verification. A malformed, mismatched, contradictory, or interrupted post-dispatch receipt is an
unknown outcome. Recovery must not turn that uncertainty into automatic retry authority.

`changeId` is minted for the request and carried through preview, approval, execution, and receipt.
It is a correlation key, not a universal host idempotency guarantee. Supported cell writes use
snapshot/readback/recovery rules described in [compute and recovery](COMPUTE-RECOVERY.md).

## Capability authority

Each bridge defines one executable handler table. `createBridgeDispatch` derives the immutable
handled-kind list from its own keys and validates requests before host access. Handwritten mirror
lists must not be reintroduced.

The effective `CapabilityManifest` determines what a particular session can request. Handler
presence determines what the host adapter can execute. The capability descriptor registry describes
discovery and policy metadata; it cannot grant a missing capability or install a handler.

Closure tests reject advertised actuations without handlers and advertised reads without ports.
Read-only runtime inspection does not grant write authority. Generated CLI preflight remains
advisory: runtime admission, approval, gates, and bridge validation still execute.

## Grammar, programs, and skills

The existing `cmd` and `plan` delimiters remain. Commands, expressions, macros, analysis bindings,
and deterministic SDK programs converge on the existing compiler and effect boundary. A pure
pipeline cannot contain an effect. Effect arguments are resolved during dry-run before approval;
macro expansion cannot bypass that process, inject extra command lines, or escape execution budgets.

The human intent tier remains `ask`, `summarize`, `explain`, `rewrite`, `review`, `draft`, and `notes`.
Scope describes where; grounding describes the evidence. Write/annotation routes enter the command
approval path rather than ordinary chat. Quick actions and palette options are capability-scoped.

The TypeScript planner parser owns supported planner keywords. Cross-surface guidance uses supported
`step`/`clarify` lines; Python-only `workflow`, `source`, `target`, `phase`, and `handoff` directives are
not accepted planner syntax. See [the planner skill](../skill/m365-command-planner/SKILL.md).

Generate CLI manifests with `bun run skills:generate`, then run `bun run skills:check`. The Python
loader rejects absent, corrupt, incompatible, or incomplete generated metadata rather than
substituting empty capability sets. Unsupported schema checks stop metadata generation. The
standalone bundles carry the shared loader; archive validation checks exact source parity and
reproducible content. See [skill maintenance](../skill/README.md).

## Ownership and attribution

`CommandContextSession` owns command observation retention, encoding, mode policy, and reference
lifetime. `AssistSession` owns execution and provider delivery. A context projection is data, not
a callback into host execution and not an approval grant.

`PanelController` acquires one execution owner before publishing busy state. Only that owner can
release the operation, synchronize history, and drain queued turns. Stream routes share their event
driver; planner handoffs retain explicit ownership transfer. Proposals pass cancellation through
the SDK boundary and keep returned receipts after non-abortable host dispatch.

Provenance belongs to the effect's originating turn or proposal. Omitted proposal attribution is
recorded as missing instead of borrowing an unrelated chat's provenance. Word and Excel share the
record serializer while retaining their host-specific storage encodings and readback. Other bridges
report unsupported durable persistence explicitly. Pane history retains verification and recovery
flags and returns detached records so consumers cannot rewrite its stored audit state.

Extension interfaces remain typed in-process contracts. See [runtime extensions](RUNTIME-EXTENSIONS.md)
for phase payloads, allowed decisions, cancellation, and observer behavior.
