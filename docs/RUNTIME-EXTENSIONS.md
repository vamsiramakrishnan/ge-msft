# Runtime extensions

Runtime extensions add context providers, checks, and event reactions without modifying the command
parser, controller, or individual Office bridges. The production task pane installs them in
`packages/web-shell/src/runtime-extensions.ts`. The Outlook function-command runtime installs the
same definitions in its **own registry**. These runtimes do not share memory.

The shipped bundle installs an outcome verifier, an active-message suggestion, and a meeting-ended
suggestion. The pane now owns an `Orchestrator` that routes bridge events into working context,
refreshes context chips when idle, and displays suggestions and bounded execution diagnostics.

## Receive a request and supply context

Add a trusted, compiled `RuntimeExtension` to `APPLICATION_EXTENSIONS`. `setup` is synchronous;
asynchronous work belongs in a handler and must honor its abort signal. Registration is atomic,
namespaced by extension ID, and reversible through the installation's disposer.

```ts
import type { ResolvedContext } from '@ge/contracts';
import type { RuntimeExtension } from '@ge/runtime';

export function projectContextExtension(
  loadFacts: (text: string, signal: AbortSignal) => Promise<ResolvedContext[]>,
): RuntimeExtension {
  return {
    id: 'tenant.project-context',
    setup(api) {
      api.on({
        id: 'receive',
        on: 'message:received',
        mode: 'guard',
        timeoutMs: 2000,
        async handle({ mode, text }, { signal }) {
          if (mode === 'proposal') return;
          const entries = await loadFacts(text, signal);
          return { kind: 'context', entries };
        },
      });
    },
  };
}
```

`message:received` means a task request entering `ask`, `runCommands`, `runCommandProgram`, `plan`,
or explicit proposal `apply`. It runs before that task's model or host work. Its `mode` distinguishes
these paths (`chat`, `command`, `program`, `planner`, `proposal`); proposals have empty request text.
It runs once per task, not once per model turn. The SDK's explicit internal `commit('prime')` is a
context-only model turn, not a new user request.

Returned context is validated with `ResolvedContextSchema`, assigned task-specific IDs, and passed
through the existing untrusted-data framing. The combined receive-phase budget is 16 entries and
64 KiB. Entries are removed from local session context when the task finishes, fails, or is cancelled;
content already sent to Gemini can remain in the remote conversation history. A provider must read
only sources authorized for the signed-in user. Hooks receive no extra Microsoft or Google permissions.

## React to an incoming host message

Request receipt and host-message activity are separate events. Use a trigger to produce a visible
suggestion, or a `host:event` observer for application-owned work that does not need an outcome:

```ts
import type { RuntimeExtension } from '@ge/runtime';

export const mailReview: RuntimeExtension = {
  id: 'tenant.mail-review',
  setup(api) {
    api.trigger({
      id: 'received',
      on: 'mail-received',
      handle: () => ({
        kind: 'suggest',
        title: 'Review commitments',
        query: '/ask @this Identify explicit commitments and unresolved questions.',
      }),
    });
    api.on({
      id: 'observe-host',
      on: 'host:event',
      mode: 'observe',
      handle({ event }, { signal }) {
        if (event.type !== 'mail-received' || signal.aborted) return;
        // Update application-owned state using event.id; do not copy message bodies into logs.
      },
    });
  },
};
```

The current Outlook bridge emits `mail-received` when a received message becomes active through
`ItemChanged` in the pane. This is **not a background mailbox-delivery subscription**. The new
`connectPanelRuntime(...).receive(event)` entry point lets a trusted adapter publish additional
`HostEvent`s while the connection is running. A future Graph notification adapter must authenticate
and validate its external payload before calling it. No Graph webhook, polling service, or Teams bot
is installed by this change.

The pane turns `automate` trigger outcomes into explicit suggestion chips too. Accepting a chip uses
the normal task router and approval flow. Merely receiving an event makes no model call in production:
host changes fold into the next request. SDK callers retain the existing `primeOnHostEvent` default;
production explicitly sets it to `false`.

## Hook phases and authority

| Phase | Payload | Allowed behavior |
| --- | --- | --- |
| `message:received` | Task mode and request text | Observe, block, or supply task context |
| `model:request` | Query text and model route | Observe or block before the model request |
| `model:event` | One typed SSE event | Observe or block before that event reaches its consumer |
| `model:response` | Accumulated answer and route | Observe or block before command parsing |
| `tool:before` | Operation name and arguments | Observe or block a runtime read/workspace operation |
| `tool:after` | Operation name and result | Observe only |
| `plan:ready` | Resolved actuation requests | Observe or block before plan approval |
| `effect:before` | Approved, provenance-stamped request | Observe or block before the trigger gate and host write |
| `effect:after` | Request and actual result | Observe only, including failed and skipped effects |
| `task:verify` | Task outcome and effect receipts | Observe or block successful completion |
| `task:finished` | Final status and effect receipts | Observe only, including cancellation |
| `host:event` | Debounced/admitted host event | Observe only |

Returning `continue` never grants write approval. Hooks cannot rewrite an approved request. The
existing plan/write approval and actuation trigger gate remain authoritative. A task verifier may
reject completion **after some effects have applied**; that neither rolls back nor retries writes.

`tool:before` covers session context list/resolve/snapshot/search, command read intents, and workspace
operations. Names are the read intent (such as `read`), `context:list`, `context:resolve`,
`context:snapshot`, `context:search`, or `workspace:<operation>`. Tool results retain their existing
shapes; read failures may be corrective `{error: ...}` values rather than thrown exceptions.
`toolCalls` counts wrapper invocations and can include nested context operations.

Workspace `/shared` operations use `workspace:share` plus the existing separate human share approval;
they are not `ActuationRequest`s and do not appear in the ledger's `effects` array. Their result and
existing share-provenance record remain available. Context tray discovery, authentication, direct
file upload, custom functions, and direct SDK client calls are not all routed through `AssistSession`
and are not claimed to be globally intercepted. The Outlook send gate uses `api.trigger` with
`on: 'mail-send'`, not an `effect:before` hook in another browser runtime.

`model:event` can withhold the current event, but cannot retract earlier streamed tokens.
`model:response` runs after text has streamed. Use it to validate commands before execution; it is not
whole-answer pre-display moderation. Model policy blocks and stream errors halt the task before
command parsing. Engine-side policy screening remains separate.

## Verify outcomes

```ts
import type { RuntimeExtension } from '@ge/runtime';

export const requireReversibleCells: RuntimeExtension = {
  id: 'tenant.cell-receipts',
  setup(api) {
    api.on({
      id: 'verify',
      on: 'task:verify',
      mode: 'guard',
      handle({ outcome }) {
        const missing = outcome.effects.some(
          (effect) => effect.ok && effect.kind === 'write-cells' && !effect.inverse,
        );
        if (missing) {
          return {
            kind: 'block',
            reason: 'A cell change applied without an undo receipt. Review the workbook before continuing.',
          };
        }
      },
    });
  },
};
```

The installed `core.outcomes` verifier rejects incomplete executions: failed/skipped effects,
exhausted or capped loops, and final program results containing errors. Intermediate model turns can
still recover from read/parse errors. It does not prove formula correctness, factual grounding, or
business invariants. Implement those checks explicitly using `task:verify` and authorized readback.

The final `done` event is delivered only after verification succeeds. Per-model-turn SSE `done`
events are not forwarded as task completion in command mode.

`task:finished` runs once when the consumer drains or closes the task iterator. Consumers must drain
async generators or call `return()`; an abandoned iterator cannot finish itself. One mutable task
owns a session at a time. Cancellation is checked after approval and before host actuation; an Office
write already in flight may still land, and its returned receipt must be preserved.

## Execution guarantees and limits

- Hooks run in descending priority, then registration order. A dispatch snapshots registrations;
  registering/unregistering during a handler affects later dispatches.
- Each handler receives an independently cloned, deeply frozen payload. This protects runtime data
  from accidental mutation. Extensions are trusted JavaScript compiled into the app, **not a sandbox**.
  Never register executable code from documents, model output, or remote configuration.
- Default hook timeout is 750 ms; configurable range is 1–10,000 ms with a 10-second total phase
  budget. Timeouts and cancellation abort the supplied signal and discard late results. CPU-blocking
  JavaScript cannot be preempted; extensions must avoid it.
- Guard exceptions, invalid results, and timeouts stop the operation. Observer failures are isolated
  and recorded. A broken post-write observer cannot turn an applied write into a failed write or
  trigger a retry. Trigger handlers have a 750 ms default timeout and a 5-second dispatch/gate budget;
  registered send/pre-actuation check failures block the gated operation. An empty registry allows it.
- The orchestrator serializes events, debounces selection/document events independently, and caps
  admitted work at 64 events. A full queue reports `host_event_backlog` and rejects additional ingress;
  it is not a durable event bus. Stop cancels debouncers/context work and suppresses late reactions.
  Restart is idempotent. Production emits `session-start`; stop does not promise a `session-end` event.
- `session.hooks.records()` keeps the most recent 256 metadata records. `session.executions.list()`
  keeps 100 task summaries, with task/surface/mode/status, timestamps, counts, and effect IDs/kinds/error
  codes. Neither stores prompts, source bodies, model output, exception bodies, or inverse payloads.
  Subscribers receive detached copies and cannot affect execution by throwing.
- Full effect and inverse receipts are ephemeral hook inputs. `ProvenanceStore` now retains a detached
  inverse receipt when supplied by the host. This foundation does not implement persistent replay,
  global exactly-once delivery, automatic compensation, or a user-facing undo executor.

## Validation

`hooks.test.ts` covers dispatch isolation, budgets, ordering, cancellation, and registration rollback.
`lifecycle.integration.test.ts` exercises real `AssistSession` entry points using fake network and
host adapters. Orchestrator and panel integration tests cover live wiring, stop/restart, suggestions,
and separate send-runtime registrations. Trigger hardening tests cover failed and hung gates.
Live Office and configured Entra/Gemini tenant validation remains a release gate.

## Analysis, evidence and recovery services

`AssistSession.runAnalysis(action, options)` shares task ownership, hooks and plan approval with
command execution. `message:received` also exposes the request's structured `dataStoreSpecs` to
trusted evidence providers. See [COMPUTE-RECOVERY.md](COMPUTE-RECOVERY.md) for extension ports,
source freshness, readback semantics, persistence and supported undo.
