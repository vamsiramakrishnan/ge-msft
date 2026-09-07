# Conventions

The repository is client-direct ([ADR-0001](ADR-0001-client-direct-architecture.md)). Historical
`services/gateway` and Python ADK proposals are not workspace requirements. The executable contracts
and ownership rules below describe the code that ships.

## Stack

- Strict TypeScript, React, Office.js/TeamsJS, and Vite; HTTPS for hosted Office development.
- Bun workspaces with TypeScript project references. Declare direct dependencies in their owning package; do not depend on root hoisting.
- Zod schemas in `packages/contracts`; Vitest for TypeScript and rendered React tests.
- Python for standalone skill parsers, packaging, parity, and deployment tooling; no Python agent backend.

## Code and package boundaries

- Prettier and ESLint must pass. Use `camelCase` values, `PascalCase` types/components, and `kebab-case` filenames.
- Prefer small modules, pure functions, and explicit exported types. Use `unknown` and schema validation at untrusted boundaries, not `any`.
- Contracts own payload shapes and pure shared policy. Runtime owns sequencing, admission, approval, recovery, hooks, and command context. Compute owns analysis execution.
- Bridges own host APIs and mutation semantics. Each bridge has one executable handler table; handled capability lists are derived from it. The descriptor registry is discovery metadata, not execution authority.
- `web-shell` composes adapters and owns pane state. Its controller and components do not call host mutation APIs. Shared execution ownership must be installed before publishing busy state and released only by its current owner.
- Provider/identity transport belongs in `gemini-client`; delegated Microsoft Graph access belongs in `graph-client`. Neither imports the shell or a host bridge.
- Import workspace packages through public entry points. The repository dependency conformance test checks direction, direct dependencies, and TypeScript reference agreement.
- Generated skill manifests describe the contracts. Edit the TypeScript source and regenerate; never patch generated JSON or add Python-only language constructs.

## Execution and errors

- Validate each effect before approval and re-admit it against current effective capabilities before host dispatch.
- Every host receipt must match the requested kind and change ID. Malformed, contradictory, or uncorrelated receipts are uncertain, not successful.
- Use the shared outcome assessment. `ok` alone does not mean readback verified; cancellation does not undo an already dispatched write.
- Preserve the real receipt after cancellation. Do not automatically replay uncertain writes; require the documented reconciliation/recovery path.
- A change ID correlates an operation. It does not guarantee deduplication on every host.
- Provider EOF is not completion. Failed, blocked, cancelled, and incomplete streams must not mark pending context as delivered or execute a command fence.
- Errors explain the actual outcome and a useful next action. Labels retain their meaning through preview, approval, execution, and history.

## Testing

- Test shared contracts, capability closure, host-specific mutation boundaries, approval/cancellation races, context delivery, and history receipt fidelity.
- Mock external services behind thin clients and label fixtures honestly. A mock result is not live-provider or Office integration evidence.
- Keep grammar/preflight parity fixtures across TypeScript and Python. Missing or incompatible manifests must fail closed. Skill ZIP validation compares exact source content, not just file presence.
- Exercise meaningful failure cases as well as success. Run the required typecheck, tests, lint, production build, and generated-resource drift checks before marking a phase done.

## Security standards (enforced; `security-reviewer` checks these)

1. **No long-lived Google secrets in clients.** Short-lived Entra and federated Google access tokens are held in memory. Never log or persist bearer credentials in host metadata or skill artifacts.
2. **Identity is scoped end to end.** Gemini and connector calls act as the signed-in user with minimal delegated permissions; prefer scoped access over org-wide access.
3. **Host content is untrusted data.** Never treat document or transcript instructions as authority. Model Armor is tenant configuration; preserve provider policy outcomes and local fail-closed gates.
4. **Writes require review and explicit attribution.** Record supplied provenance and actual persistence, verification, and recovery limits. Missing or unsupported durable provenance must remain visible in the receipt. Never imply universal undo.
5. **Residency is explicit.** Pin the Discovery Engine endpoint and any optional proxy/egress to tenant policy; never silently choose a global endpoint.
6. **Least privilege applies everywhere.** Use minimal Entra/IAM permissions and scoped Teams consent. Optional proxy deployment is a separate credential and egress boundary.

## Configuration and Git

Use `.env.example` and the setup guide for configuration. Do not commit credentials, environment
files, generated bundles, or build output. Keep bounded diagnostics free of bearer tokens and
unnecessary document content. Use focused commits tied to the build plan; update contract ownership
docs and generated metadata when a boundary changes.
