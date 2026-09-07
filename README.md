# Gemini Enterprise for Microsoft 365

**Use Gemini Enterprise from the Microsoft 365 document or conversation you are already working in.**

A Microsoft 365 add-in that exposes Gemini Enterprise inside Word, Excel, PowerPoint, OneNote, Outlook, and Teams.

The browser add-in exchanges the signed-in user's Entra token for a short-lived Google token through Workforce Identity Federation, then calls Gemini Enterprise / Discovery Engine as that user. The current architecture is client-direct; there is no credential-holding application gateway in the normal request path.

Document changes go through surface-specific bridges so the runtime can attach provenance and use each host's native mutation model.

The Excel **Data workbench** now supports versioned range snapshots, local SQL, exact-decimal
reconciliation, finding chips, reviewed writeback, readback verification, and durable cell undo.
Selected data stores also feed an evidence assembly and claim-checking pipeline. See
[the implementation and operating limits](docs/COMPUTE-RECOVERY.md).

## Choose the environment you want to test

| Goal | Start here | Successful result |
|---|---|---|
| Inspect the task pane | [Local preview](#start-locally) | The UI runs without an Office host |
| Exercise a real Office surface | [Setup guide](setup/README.md) | The add-in is sideloaded and configured for that host |
| Depend on a particular read or write | [Current status](docs/STATUS.md) | The bridge implements the capability and its persistence behavior |
| Prepare a tenant release | [Deployment choices](setup/07-deployment-methods-matrix.md) | An explicit hosting and distribution path |

Previewing the pane, authenticating to Gemini Enterprise, and changing an Office
document are separate checks. Start with the surface you need; support and
provenance persistence differ between hosts.


## Start locally

Use Bun for repository tasks:

```bash
bun install
bun run setup:doctor
bun run --filter @ge/web-shell preview
```

The preview renders the task pane without requiring an Office host. `bun run dev` opens the same preview server. Choose **Try interactive demo** to exercise source attachment, answers, and approval decisions with scripted sample data. Width and height controls reproduce narrow Office panes. No model calls or Office writes occur in the demo.

Useful setup and release commands:

| Task | Command |
| --- | --- |
| Check workstation dependencies | `bun run setup:doctor` |
| Guided setup | `bun run setup:guide` |
| Start the development tunnel | `bun run ge:dev:tunnel` |
| Build the Office package | `bun run setup:package` |
| Sideload for the current developer | `bun run sideload` |
| Local development bootstrap | `bun run bootstrap:dev` |
| Validate a release without publishing | `bun run bootstrap:release:dry-run` |
| Publish the configured release | `bun run bootstrap:release` |
| Manage Gemini Enterprise skills | `bun run ge:skills` |

`bootstrap:dev` does not upload to the tenant catalog. Release tasks use the configured production-like profile and reject placeholder identifiers, localhost origins, and example domains where the release configuration requires stable values.

## Request path

```text
Microsoft 365 host
      │
      ▼
Office / Teams webview
      │
      ├── Office.js / TeamsJS ──► document or host state
      │
      ├── Entra token
      │       │
      │       ▼
      │   Google STS / WIF
      │       │ short-lived Google token
      │       ▼
      └─────────────────────────► Gemini Enterprise / Discovery Engine
```

The hosted origin serves the web application. It is not automatically an application backend.

An optional transparent proxy can be configured for tenants whose browser policy prevents direct calls to the Discovery Engine endpoint. That proxy is a deployment choice, not the default identity architecture.

## Package architecture

```text
surface bridges
  word · excel · powerpoint · onenote · outlook · teams
                  │
                  ▼
             DocBridge
                  │
                  ▼
      runtime + web-shell
                  │
        ┌─────────┼─────────┐
        ▼         ▼         ▼
     content   gemini-    graph-
               client     client
        │         │         │
        └─────────┼─────────┘
                  ▼
              contracts
```

The main packages have separate responsibilities:

| Package | Responsibility |
| --- | --- |
| `contracts` | Shared TypeScript types and Zod schemas for document state, commands, plans, capabilities, provenance, and actuation |
| `content` | Converts host document state into typed blocks and bounded context |
| `gemini-client` | Workforce Identity Federation and Gemini Enterprise / Discovery Engine requests |
| `graph-client` | Delegated Microsoft Graph reads |
| `triggers` | Host events, debounce, and the actuation gate |
| `runtime` | Assist session, command parsing, composition, planning, and execution |
| `web-shell` | Authentication wiring, task-pane state, React UI, and standalone preview |
| `bridge-*` / `teams` | Host-specific implementations of `DocBridge` |

The architecture decisions under `docs/ADR-*.md` describe the current design. [`docs/STATUS.md`](docs/STATUS.md) records what is implemented now. Older design notes may describe superseded gateway-based approaches.

## Document context

The runtime treats the active Office document as an addressable environment rather than serializing the entire file into every model request.

A bounded document snapshot carries current structure. The model can request narrower reads such as search, outline, or addressed content through the bridge.

Host content is wrapped as untrusted data before it is added to model context. The surrounding application still depends on Gemini Enterprise engine configuration and its configured controls; wrapping content is not a complete prompt-injection defense by itself.

## Command protocol

Discovery Engine `streamAssist` does not expose the same native function-calling contract as a standard tool-calling model. This project therefore uses a small command language for host operations.

The model emits command or plan blocks. The runtime parses and validates them, then compiles actuating commands into typed `ActuationRequest` values.

```text
model output
    │
    ▼
command / plan parser
    │
    ▼
typed values and ActuationRequest
    │
    ▼
dry run / approval
    │
    ▼
DocBridge.actuate(...)
```

Pure transformations and host mutations are kept separate. Reads and transformations can be evaluated during a dry run. Effects are collected for approval before execution.

## Capability closure

Each Office surface supports a different set of reads and mutations.

The contracts package compares the declared capability manifest with implemented bridge handlers, read ports, and command verbs. Conformance tests fail when a surface advertises a capability that its implementation does not handle.

This check covers the capabilities represented by those registries and tests. It is not a proof that arbitrary host behavior cannot diverge outside the checked surface.

## Work from the task pane

Start with the current document. Attach the sources that matter. Choose a task or describe the result you need.

| Control | What it does |
| --- | --- |
| Context chips | Attach, remove, and locate host content without opening a dialog. The controller owns the attachment state. |
| Action library | Search 47 actions across six surfaces. Filter by Ask, Review, or Change; pin the ones you use. Only supported actions appear. |
| Task intent and scope | Choose what to do and where to do it. Slash commands remain available. |
| Request sources | Pick connected data stores by name. Their resource IDs become structured Gemini grounding for this request. |
| Response controls | Request a comparison table, decision brief, checklist, or concise bullets; choose a writing style. |
| Answer controls | Copy a completed answer or prepare an evidence check, decision brief, or action checklist as an editable follow-up. |
| Change review | Inspect the number of changes, affected targets, exact commands, and available before/after values before approving. |

Open **Actions** with **Ctrl+K / ⌘K**. **Enter** submits; **Shift+Enter** adds a line. Input-method composition does not accidentally submit.

The catalog includes workflows for decision reviews in Word, reconciliation and chart analysis in Excel, evidence checks and decision decks in PowerPoint, research briefs in OneNote, commitments and reply drafts in Outlook, and decision logs in Teams. These workflows use the existing bridges and approval path. They do not add permissions or claim host capabilities that are unavailable.

See the [workspace guide](docs/WORKSPACE.md) for the complete interaction flow and limits. Gemini Enterprise [skills](skill/) carry the underlying command grammar.

## Actuation and provenance

Writes pass through the actuation gate.

For plans, the runtime evaluates reads and pure transformations first, resolves the pending effects, and presents that effect set for approval before mutation.

Provenance records include the acting agent, source references, identity, timestamp, and content hash. Persistence differs by Office surface. Word and Excel have host-backed provenance storage in the current implementation; other surfaces have different or incomplete persistence paths. Check [`docs/STATUS.md`](docs/STATUS.md) before depending on a particular host's persistence behavior.

Excel formula writes also pass through the project's formula-safety check before actuation.

Word writes can use content anchors and re-resolve the target before applying a change. When the anchor no longer resolves, the bridge can refuse or degrade the action rather than writing to the stale position.

## Identity boundary

The normal browser path holds the user's short-lived Entra token and the Google token derived from it in memory.

No long-lived Google service-account secret is shipped to the add-in.

This design still depends on correct Entra configuration, Workforce Identity Federation policy, browser token handling, and Gemini Enterprise IAM. Client-direct does not remove those controls; it changes where credential exchange and API calls occur.

## Research context

The application can compose context from the working document, federated Microsoft 365 sources, and configured research/notebook material. Each source type has its own retrieval and authorization path.

Do not treat the phrase "research unit" as an authorization boundary. Access is determined by the underlying Microsoft and Google identity/configuration paths.

## Side-pane previews

The checked-in GIFs are generated from the task-pane preview harness:

```bash
bun run docs:gifs
```

On a new Linux workstation, install the Playwright browser dependencies once with:

```bash
bun run docs:gifs:install
```

The preview harness verifies the UI path it exercises. It does not substitute for testing inside every Office host.

## Repository layout

```text
packages/
  contracts/       shared schemas, command grammar, plans, capabilities
  content/         host document normalization and chunking
  gemini-client/   WIF + Discovery Engine client
  graph-client/    delegated Microsoft Graph reads
  triggers/        host events and mutation gate
  runtime/         assist loop, command compiler, plan executor
  web-shell/       React task pane and standalone preview
  bridge-*/        Office host adapters

skill/             Gemini Enterprise skills
setup/             workstation, sideload, release and hosting guides
docs/              ADRs, API notes, mockups, status and architecture docs
```

## Read next

- [`docs/STATUS.md`](docs/STATUS.md): implemented surface and known gaps
- [Runtime extensions](docs/RUNTIME-EXTENSIONS.md): message hooks, context providers, event reactions, and outcome verification
- [`setup/README.md`](setup/README.md): setup path
- [`setup/07-deployment-methods-matrix.md`](setup/07-deployment-methods-matrix.md): deployment choices
- [`setup/08-hosting-origin-and-release.md`](setup/08-hosting-origin-and-release.md): hosting boundary
- `docs/ADR-0001-*`: identity architecture
- `docs/ADR-0003-*`: document-as-environment model
- `docs/ADR-0004-*`: command protocol
- `docs/ADR-0005-*`: composition, plans, and skills
- `docs/ADR-0006-*`: capability closure

## Boundaries

- Direct browser calls do not remove the need for IAM, tenant policy, and token-handling controls.
- A declared capability is only valid for a surface when the bridge implements and passes its conformance checks.
- A dry run prevents host mutation during that phase; it cannot predict unrelated changes another user or process makes before approval.
- Provenance persistence is surface-specific. Check current status before assuming a record survives save/reopen on every host.
- Host content is untrusted input. Wrapping and engine-side screening reduce exposure but do not establish that prompt injection is impossible.

## Development

```bash
bun install
bun run setup:doctor
bun test
```

Use the ADRs for architecture decisions and `docs/STATUS.md` for the current implementation boundary.

## License

See [LICENSE](LICENSE).
