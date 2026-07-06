# ADR-0009 — Tenant bootstrap, ring configuration, and bounded deployment envelopes

**Status:** Proposed (2026-06-27); first slice implemented (2026-07-06) — same-origin static config with a `?cfg=` selector and fail-closed validation ships in `packages/web-shell` (see "Implemented slice" below); the SharePoint/ring resolver remains proposed · extends ADR-0001 (client-direct architecture), ADR-0002 (capability model), ADR-0006 (capability closure), and ADR-0008 (Surface Commander algebra). Scope: deployment bootstrap, tenant configuration, rollout rings, and the trust boundary between Microsoft identity, SharePoint configuration, and Google Workforce Identity Federation.

## Implemented slice (2026-07-06)

The shell now resolves a **runtime tenant config** at boot (`resolveRuntimeEnv` in
`packages/web-shell/src/taskpane/config.ts`, awaited in `main.tsx` before any config accessor):

- A `?cfg=<name>` query parameter (validated against `^[a-z0-9][a-z0-9-]{0,63}$`; invalid values
  are ignored, never interpolated) selects `${origin}/config/<name>.json`, with
  `${origin}/config/default.json` as the fallback. Candidates are **same-origin only** by
  construction; the fetch uses `cache: 'no-store'` and a ~4s timeout.
- A fetched document must be a flat JSON object whose keys are exactly the known public `VITE_*`
  config keys with string values. Unknown keys, non-string values, or anything matching the
  browser secrets denylist rejects the **whole** document (fail closed to the build-time env).
- Accepted values overlay the build-time env, and the merged result still runs through the
  unchanged `parseEnv` validation (required keys, HTTPS-only, authority allowlist,
  unknown-prod-key rejection). The runtime document can select values; it can never widen the
  schema, mint a permission, or carry a secret.

This lets one centrally hosted bundle serve multiple tenants (per-tenant coordinates ride on the
manifest content URL's `?cfg=`), without yet building the SharePoint store, app-role rings, or
locator service described below.

Recorded residual risk (per security review of the slice): the runtime document may legally
override any known public key, including `VITE_PROXY_URL` (any HTTPS host), the WIF pool/provider,
and the GCP project — so write access to `/config/*.json` on the add-in origin is equivalent to
retargeting the add-in's egress within the schema. Treat config publishing with the same change
control as bundle publishing, never host weaker configs (dev proxies, test pools) on the
production origin (`?cfg=` lets a crafted deep link select any published name), and before
production tenants either narrow the runtime-overridable key set, allowlist the production proxy
host (the `proxyRef` design below), or integrity-protect the document. WIF/IAM conditions (§7)
remain the token-layer backstop that makes retargeted coordinates non-useful.

## Context

The add-in should not require manifest edits for ordinary configuration changes. The manifest should carry stable extension points and a stable web-shell URL; mutable product routing should live outside the manifest so tenant admins can roll features, engines, and surfaces without repackaging the add-in.

However, the first design instinct — "provision the maximal Microsoft Graph scopes we might ever need and freeze the manifest" — is wrong. A Word/Excel alpha should not ask a customer to consent to broad future scopes such as mailbox write or tenant-wide SharePoint write. Manifest and Entra consent form a security envelope, not a convenience cache.

We need four different controls to stay separate:

- **Entra identity** answers who the signed-in user is and which rollout ring they are in.
- **SharePoint configuration** answers what that ring receives.
- **The build-time release profile** answers which product envelope this artifact was shipped to support.
- **Google WIF/IAM** answers which Google resources a federated token can actually use.

Configuration must never create trust. It may select, route, reduce, or disable inside a pre-existing envelope; it must not mint Microsoft permissions, select arbitrary egress origins, or route user content to unsanctioned Google resources.

## Decision

Adopt a bounded tenant-configuration bootstrap model:

```text
manifest / app registration
  = stable web-shell URL
  + stable host extension points
  + bounded Microsoft Graph scope ceiling for the release profile

Entra app-role values
  = rollout ring identity

SharePoint config
  = tenant-owned ring policy and routing inside the release envelope

release profile
  = build-time product envelope, not tenant-editable config

WIF provider + Google IAM
  = token-layer enforcement of sanctioned tenant/project/resource access
```

The invariant is:

```text
manifest scopes are an upper bound;
tenant config can only disable, route, or select within that bound;
tenant config can never create a permission, proxy origin, release profile, or Google trust boundary.
```

## Detailed decisions

### 1. Scope ceiling is bounded by release profile

Provision only the maximum Microsoft Graph scopes justified by the current release train or deployment profile.

For `internal-alpha-word-excel`, that means a narrow ceiling such as:

```text
openid
profile
User.Read
Sites.Selected or a narrower SelectedOperations scope for the config resource
```

It does **not** mean consenting to future estate-write or mail-write scopes. A later estate-write preview should be a new release profile, manifest package, consent event, and release gate.

### 2. Rings are Entra app-role values, not group IDs

Rollout rings are represented by app-role values emitted in the token for this app, for example:

```text
GE.Ring.Prod
GE.Ring.Pilot
GE.Ring.Internal
```

The runtime reads ring membership from the app's identity claims / ID-token path, not from a Microsoft Graph access token. A Graph access token is issued for Microsoft Graph and is not the authority for this app's app-role assignments.

Ring config keys therefore use app-role values:

```jsonc
{
  "ringPrecedence": ["GE.Ring.Prod", "GE.Ring.Pilot", "GE.Ring.Internal"],
  "rings": {
    "GE.Ring.Pilot": { "engine": { "engineId": "pilot-engine" } },
    "GE.Ring.Internal": { "features": { "surfaceCommander": true } }
  }
}
```

Nested group membership is not assumed to flow into app-role assignment. Ring groups must be leaf groups or otherwise flattened by tenant identity governance.

### 3. SharePoint stores ring config

Tenant-editable configuration lives in a known SharePoint file or list, read via Microsoft Graph using Selected permissions.

Preferred v1:

```text
SharePoint site: /sites/GeminiEnterpriseConfig
File:            /Shared Documents/config/config.json
Permission:      Sites.Selected read grant on that site
```

Harder-but-narrower later options:

```text
Lists.SelectedOperations.Selected
ListItems.SelectedOperations.Selected
Files.SelectedOperations.Selected
```

SharePoint is the default because it gives tenant admins a familiar edit surface, version history, ownership inside the customer M365 boundary, and auditability without broad directory write/read scopes.

### 4. Tenant config schema

The full shape evolves from the bootstrap probe, but the steady-state document is conceptually:

```jsonc
{
  "schemaVersion": 1,
  "default": {
    "engine": {
      "projectNumber": "123456789",
      "location": "global",
      "engineId": "prod-engine"
    },
    "surfaces": ["word", "excel"],
    "features": {
      "surfaceCommander": true,
      "excelNativeWrites": true,
      "crossSurfacePlans": false
    },
    "proxyRef": null
  },
  "ringPrecedence": ["GE.Ring.Prod", "GE.Ring.Pilot", "GE.Ring.Internal"],
  "rings": {
    "GE.Ring.Pilot": {
      "engine": { "engineId": "pilot-engine" },
      "surfaces": ["word", "excel", "powerpoint"]
    },
    "GE.Ring.Internal": {
      "features": { "crossSurfacePlans": true }
    }
  },
  "metadata": {
    "version": 7,
    "updatedBy": "admin@example.invalid",
    "updatedAt": "2026-06-27T00:00:00Z"
  }
}
```

Unknown keys are rejected. The resolver performs:

```text
TenantConfig
  -> validate
  -> select matching roles using ringPrecedence
  -> merge default + selected rings
  -> validate resolved config
  -> intersect with release profile
  -> intersect with manifest/app consent ceiling
  -> intersect with runtime host capabilities
```

Merge semantics are explicit:

```text
objects: deep merge
arrays: replace, never concatenate
scalars: replace
null: explicit disable only where schema permits
claim order: ignored
ringPrecedence: the only source of ring order
```

The recommended precedence order is least-specific to most-specific with **last matching ring wins**:

```json
["GE.Ring.Prod", "GE.Ring.Pilot", "GE.Ring.Internal"]
```

This avoids accidental "most permissive" composites that no admin intended.

### 5. No tenant-editable `proxyUrl`

Tenant config must not contain an arbitrary `proxyUrl`. A tenant-editable URL that receives a federated bearer token is a token-exfiltration primitive; HTTPS validation does not help because `https://evil.example` is still valid HTTPS.

Config may contain only:

```jsonc
{ "proxyRef": "japac" }
```

The shell resolves `proxyRef` through a build-time allowlist or same-origin policy:

```ts
const ALLOWED_PROXY_BY_REF = {
  default: "https://ge.example.invalid/proxy",
  japac: "https://japac.ge.example.invalid/proxy"
} as const;
```

A missing or unknown `proxyRef` fails closed for proxy mode.

### 6. Release profile is build-time

The release profile is a compiled constant of the shipped artifact. It cannot come from SharePoint config.

Example:

```ts
export const RELEASE_PROFILE = {
  id: "internal-alpha-word-excel",
  surfaces: ["word", "excel"],
  graphScopeCeiling: ["openid", "profile", "User.Read", "Sites.Selected"],
  features: {
    outlook: false,
    teams: false,
    crossSurfacePlans: false,
    estateWrites: false
  }
} as const;
```

The token proves Microsoft API permission. The release profile proves product packaging, tested surface area, diagnostics posture, provenance policy, and allowed config providers.

### 7. WIF/IAM enforces config non-escalation at the token layer

Zod validation is not the strongest control. Google IAM must also make bad config non-useful.

The WIF provider maps immutable Entra claims into Google attributes, for example:

```text
assertion.tid -> attribute.tenant_id
assertion.oid -> google.subject
assertion.aud or azp -> attribute.client_id
```

The provider condition constrains accepted tokens to the expected tenant and app:

```text
attribute.tenant_id == "<tenant-guid>"
&& attribute.client_id == "<expected-entra-app-id>"
```

IAM then grants the workforce principal set only the minimum Discovery Engine / Gemini Enterprise permissions required on the sanctioned Google project or narrower supported resource:

```text
principalSet://iam.googleapis.com/locations/global/workforcePools/<pool-id>/attribute.tenant_id/<tenant-guid>
  -> minimum required role on projects/<sanctioned-project>
```

This closes two risks:

- `proxyUrl`-style token exfiltration is blocked by `proxyRef` plus allowlisting.
- malicious `engineId` / project routing becomes a content-exfiltration attempt that fails against Google IAM or the engine allowlist.

If the relevant Google API only supports project-level IAM, the engine allowlist remains mandatory defense-in-depth inside the sanctioned project.

### 8. Bootstrap modes

Support three modes over time.

#### Mode A — well-known SharePoint convention

Good for single-tenant and internal alpha installs:

```text
https://{tenant}.sharepoint.com/sites/GeminiEnterpriseConfig
/Shared Documents/config/config.json
```

Simple, but not universal across multi-geo SharePoint or customers with strict site naming policies.

#### Mode B — Tenant Config Locator Service

For real multi-tenant rollout, add a thin service:

```text
tenantId -> config locator pointer
```

It stores only non-secret pointers:

```json
{
  "schemaVersion": 1,
  "tenantId": "00000000-0000-0000-0000-000000000000",
  "provider": "sharepoint",
  "sharePoint": {
    "hostname": "contoso.sharepoint.com",
    "siteId": "contoso.sharepoint.com,abc,def",
    "driveId": "b!xyz",
    "itemId": "01ABCDEF..."
  }
}
```

It must not hold user tokens, refresh tokens, Google credentials, engine IDs, feature flags, prompts, or document content.

This is a narrow server exception to ADR-0001. It is not the old gateway. It breaks the chicken-and-egg where discovering the config site generically would otherwise tempt broad Graph discovery scopes.

#### Mode C — GCP config provider

Optionally support GCS/Firestore config for Google-first customers. This is not the default for M365-first tenants because it weakens the Microsoft admin UX and moves config governance out of SharePoint.

## First implementation slice

The first slice must be a live-host bootstrap probe, not the full ring system:

```text
one known SharePoint file
one field: engineId
one app role: GE.Ring.Probe
one host path: Word/Excel task pane
one diagnostic card
composeSession consumes fetched engineId
```

No rings, no locator, no deep merge, no proxyRef, no feature flags, no GCP provider.

The probe answers two runtime questions that fake-host tests cannot prove:

1. Does the app-role value appear in the claims available to the Office-hosted add-in through the NAA/MSAL path?
2. Can the signed-in user acquire a Graph token and read the selected SharePoint config file through the intended Selected-permissions boundary?

## Consequences

Positive:

- Config moves at SharePoint edit speed, not manifest deployment speed.
- Rings do not require `GroupMember.Read.All` or raw group graph exposure.
- Config changes do not re-trigger consent.
- The manifest consent surface is reviewable and bounded by release profile.
- A compromised config document cannot create Microsoft or Google trust.
- Multi-tenant bootstrap can later avoid broad Graph discovery scopes through a non-secret locator.

Costs:

- Initial tenant setup requires app roles and a SharePoint Selected permission grant.
- Group nesting is not supported for ring membership unless customers flatten groups.
- Dynamic configuration adds a new setup/error surface that must be diagnosable.
- Multi-tenant bootstrap eventually requires a small locator service, which is a deliberate exception to a purely client-direct deployment.
- WIF/IAM must be configured carefully; code validation alone is insufficient.

## Rejected alternatives

### Maximal-lifetime Graph scopes in the manifest

Rejected. Broad future scopes such as `Mail.ReadWrite` or tenant-wide SharePoint write permissions create a consent screen a customer security team should reject for a Word/Excel alpha.

### Raw Entra group claims for rings

Rejected. Group claims expose more tenant graph than needed and can hit overage limits. App-role values are app-specific, smaller, and auditable.

### Ring keys as group IDs

Rejected. The runtime consumes app-role values, not group IDs. Group IDs belong in tenant setup documentation, not runtime policy resolution.

### Entra open extensions as the config store

Rejected for v1. They lack a friendly admin edit surface and push customers toward Graph/PowerShell with directory-centric permissions.

### Tenant-editable proxy URLs

Rejected. Arbitrary proxy origins can exfiltrate federated bearer tokens.

### Release profile in tenant config

Rejected. A compromised or mistaken SharePoint config must not switch the product envelope.

## Validation requirements

Before enabling this for a tenant alpha:

- Sideload Word and Excel against a real tenant.
- Confirm `GE.Ring.Probe` appears from the app identity claims, not by decoding a Graph access token.
- Confirm the selected SharePoint config read succeeds with no broad SharePoint read scope.
- Confirm missing role, missing permission, 403, 404, invalid JSON, and invalid schema all produce typed setup errors.
- Confirm diagnostics redact tokens and raw JWTs.
- Confirm `composeSession` uses the fetched `engineId`.
- Confirm WIF provider conditions and IAM binding block wrong tenant, wrong app, wrong project, and unsanctioned engine routing.
