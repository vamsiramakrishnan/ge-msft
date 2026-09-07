# ADR-0013: Bounded command programs and verified completion

Date: 2026-09-07
Status: Implemented; live Office/provider measurements remain a release gate.

The expensive unit in the command loop is a model round trip. Copying generated artifact IDs back
through a model, repeating the complete grammar, dumping tables, and asking the model to say `done`
all increase work without increasing execution capability.

We retain the established `cmd` fence and typed command payloads. A new task-local artifact binding
lets one program capture, derive, review, apply and verify a result. `finish when=verified` asks the
runtime to terminate after verification; it cannot assert that verification succeeded. The SDK
compiles a versioned `AnalysisProgram` into the same execution path with no model call.

Verified programs require one complete command frame. Before execution, the runtime expands macros,
checks command/write budgets and terminal ordering, and rejects legacy `done` and unsupported shares.
Read/derivation failures discard pending writes. Host approval, freshness checks, recovery checkpoints,
readback verification and lifecycle hooks retain their existing authority. An incomplete, rejected,
cancelled or failed run cannot produce a successful completion event. This is not a transaction:
earlier writes can have landed before a later effect fails, and nothing is replayed automatically.

Artifact references substitute only in typed artifact fields and admitted SQL table positions.
They cannot alter host targets, quoted SQL strings or arbitrary JSON fields. The existing constrained
DuckDB engine validates the resolved SQL; no general-purpose model code execution is introduced.

Discovery is selective: a compact bootstrap advertises the active surface, core rules and relevant
command cards. Exact syntax remains available through `help <verb>` and `help full`. Large tool
results stay in a bounded, task-scoped store, with projected inspection and explicit continuation.
References confer no new read or write authority. Unchanged document structure is abbreviated only
within the same successful provider conversation; every turn still captures fresh host state.

The performance contract measures model calls, approvals and UTF-8 query/result bytes. These are
reproducible local metrics, not a claim about provider tokens, billing or end-to-end latency.
See [COMMAND-PERFORMANCE.md](COMMAND-PERFORMANCE.md) for examples, measurements and boundaries.

Discovery Engine API versions differ: the public v1alpha `streamAssist` contract documents an
explicit `isSessionLess` flag; the v1beta reference does not. The SDK exposes the direct v1alpha flag
as an opt-in. The application uses it by default for commands and planning, while normal chat retains
its session. Every independent command request carries the full bounded task/program/result context
and a fresh document snapshot. Overflow and session-bound uploads stop explicitly, with no silent
stateful fallback. Conversation mode remains an explicit compatibility option. Tenant behavior still
needs live validation. Omitting `session` alone creates a session.
