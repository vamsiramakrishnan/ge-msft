# Bootstrap live validation — step-one probe

This is the smallest live-host validation slice for ADR-0009.

It intentionally validates only two deployment assumptions:

1. The Office-hosted add-in can see this application's Entra app-role values through the NAA/MSAL identity path.
2. The same add-in can read one known SharePoint config file through Microsoft Graph using the intended Selected-permissions boundary.

Do not bundle rings, a locator service, proxy selection, feature flags, dynamic capability rollout, or GCP config into this first probe.

## Non-goals

This probe does **not** implement:

- ring config merging
- multi-ring precedence
- Tenant Config Locator Service
- GCP/Firestore/GCS config provider
- `proxyRef`
- feature flags
- dynamic rollout
- cross-surface plans
- estate writes
- public release readiness

## Probe contract

### Known app role

```text
GE.Ring.Probe
```

Assign either the test user directly or a **leaf** security group directly to this app role. Nested group membership is not accepted as proof; app assignment does not cascade to nested groups.

### Known SharePoint config file

Example convention:

```text
Site: /sites/GeminiEnterpriseConfig
File: /Shared Documents/config/config.json
```

Initial payload:

```json
{
  "engineId": "pilot-engine"
}
```

The schema must reject all unknown fields. This is a one-field bootstrap probe, not the final tenant config schema.

### Known Graph permission envelope

Use the narrowest Selected permission practical for the test environment. For a simple first pass, `Sites.Selected` with a read grant on the config site is acceptable.

The test must not rely on broad discovery permissions such as `Sites.Read.All` to find the site. Site/file location is configured out-of-band for the probe.

## Runtime flow

```text
Task pane initializes in Word or Excel
  -> NAA/MSAL account is available
  -> acquire/read app identity claims
  -> extract tenant id, user object id, roles[]
  -> verify GE.Ring.Probe presence
  -> acquire Graph token with the Selected scope
  -> GET known config file content
  -> Zod-validate { engineId }
  -> pass bootstrap config snapshot into composeSession
  -> composeSession uses fetched engineId instead of .env engineId
  -> render BootstrapDiagnostics card
```

Roles must be read from the app's identity claims / ID-token path. Do not decode a Microsoft Graph access token to discover app roles.

## Diagnostic card

Render one explicit card in the task pane:

```text
Bootstrap
  Host: Word | Excel | unknown
  Tenant ID: present | missing
  User object ID: present | missing
  roles claim: present | missing
  GE.Ring.Probe: present | missing
  Graph token acquired: yes | no
  Config read: yes | no
  Config ETag: <etag or none>
  engineId: <validated value or none>
```

Diagnostics must not include:

- raw JWTs
- access tokens
- refresh tokens
- subject tokens
- Google access tokens
- arbitrary SharePoint config body
- document/workbook content

## Typed failure states

Use explicit error codes:

```text
identity_unavailable
roles_claim_missing
probe_role_missing
graph_token_failed
config_forbidden
config_not_found
config_invalid_json
config_invalid_schema
config_network_error
compose_session_not_seeded
```

A missing probe role is a validation failure, not a chat message.

## Acceptance criteria

The probe passes only when all of these are true in a real sideloaded host:

- Word task pane loads and reports host identity.
- Excel task pane loads and reports host identity.
- The app-role value `GE.Ring.Probe` is visible from app identity claims for a directly assigned user.
- The app-role value `GE.Ring.Probe` is visible for a user who is a direct member of a leaf group assigned to the app role.
- A nested-group-only assignment is documented as unsupported and does not count as a pass.
- The Graph token is acquired through the add-in's NAA/MSAL path.
- The known config file is read through the Selected permission path.
- The parsed config snapshot carries the file ETag.
- `composeSession` receives the snapshot and uses `engineId` from the file.
- 403, 404, invalid JSON, invalid schema, missing role, and missing claim all show typed setup errors.
- Diagnostics and logs contain no token-like strings.

## Suggested implementation files

The exact paths may vary, but keep the boundaries clear:

```text
packages/contracts/src/tenant-bootstrap-config.ts
packages/web-shell/src/bootstrap/app-identity-claims.ts
packages/graph-client/src/sharepoint-bootstrap-config-client.ts
packages/web-shell/src/bootstrap/bootstrap-config-loader.ts
packages/web-shell/src/session/compose-session.ts
packages/web-shell/src/taskpane/components/BootstrapDiagnostics.tsx
```

The bootstrap loader should return a snapshot:

```ts
interface TenantBootstrapConfigSnapshot {
  readonly tenantId: string;
  readonly userObjectId: string;
  readonly roles: readonly string[];
  readonly source: {
    readonly provider: "sharepoint";
    readonly siteId?: string;
    readonly driveId?: string;
    readonly itemId?: string;
    readonly etag?: string;
  };
  readonly config: {
    readonly engineId: string;
  };
  readonly fetchedAt: string;
}
```

## Tests

Add offline tests for:

- schema accepts exactly `{ "engineId": "pilot-engine" }`
- schema rejects unknown keys
- claim parser extracts app-role values
- claim parser handles missing `roles`
- claim parser rejects role lookup from a Graph access token path
- Graph client sends the expected config-file request
- ETag is stored and reused in memory
- 304 reuses cache
- 403 maps to `config_forbidden`
- 404 maps to `config_not_found`
- invalid JSON maps to `config_invalid_json`
- invalid schema maps to `config_invalid_schema`
- `composeSession` uses the fetched engine ID
- diagnostics redaction catches token-like strings

## Manual live checklist

Record the following for each run:

```text
commit SHA:
manifest version:
package hash:
host: Word | Excel
platform: web | windows | mac
Office build/version:
tenant id:
test user object id:
assignment mode: direct user | direct leaf group | nested group negative test
Selected permission scope:
config site/file identifier:
result: pass | fail | blocked
failure code:
evidence screenshot/log path:
```

The result is not release evidence unless it is tied to the exact commit SHA, package hash, and manifest version under test.
