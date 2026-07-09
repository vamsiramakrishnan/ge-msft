---
title: Command Runbook
kind: reference
skill: m365-release-operator
topics: [bun, bootstrap, sideload, release, skills, entra, catalog]
load_when: Exact local commands are needed for setup, dev tunnels, sideloading, release, Entra sync, catalog upload, or skill provisioning.
---

# Command runbook

Use these commands through the local harness. Prefer the narrowest command that answers the user's
problem.

## Readiness and local setup

| Need | Command | Mutates |
| --- | --- | --- |
| Check readiness | `bun run setup:doctor` | No |
| Guided setup | `bun run setup:guide` | Maybe, asks first |
| Install prerequisites | `bun scripts/setup-cli.mjs prereqs --target all` | Local machine |
| Generate manifests/package | `bun run setup:package` | `dist/` |
| Validate skills | `bun run skills:check && python3 skill/validate_skill_bundles.py --check-zip` | Generated indexes only unless check fails |

## Dev loop

| Need | Command | Notes |
| --- | --- | --- |
| Start Vite + Cloudflare tunnel + sync dev redirect | `bun run ge:dev:tunnel` | Interactive; may update Entra SPA redirect |
| Personal unified sideload | `bun run sideload` | Uses Agents Toolkit and stores local sideload state |
| Check sideload | `bun run sideload:status` | Read-only |
| Remove sideload | `bun run sideload:uninstall` | Removes the personal sideload |
| Full dev bootstrap | `bun run bootstrap:dev` | Interactive; runs package, skills, and deploy lane |
| Dev bootstrap with sideload lane | `bun run bootstrap:dev:sideload` | Uses sideload instead of tenant deployment |

## Release flow

Use a dry-run before live release.

```bash
bun run bootstrap:release:dry-run
bun run bootstrap:release
```

For just the Entra SPA redirect step:

```bash
bun run entra:sync:release
```

The release profile must provide a stable HTTPS origin and release-specific Entra client ID. Do not
ship a Cloudflare dev tunnel URL as a production redirect.

## Gemini Enterprise skills

| Need | Command | Notes |
| --- | --- | --- |
| Rebuild and update the main widget skills | `bun run ge:skills` | Interactive; asks for widget cURL/HAR if token is stale |
| Force fresh widget token | `scripts/update-ge-widget-skills.sh --force-token-refresh` | Interactive |
| Paste cURL/HAR directly | `scripts/update-ge-widget-skills.sh --paste-curl` | Ends paste with `__GE_WIDGET_CURL_END__` |
| List visible widget skills | `scripts/update-ge-widget-skills.sh --list-only` | Also syncs non-secret env refs |
| Build this release operator zip | `(cd skill && ./build_zip.sh m365-release-operator)` | Creates `skill/m365-release-operator.zip` |
| Upload only this skill | `python3 skill/update_skills.py --api-mode widget --only m365-release-operator --create-new --yes --live` | Requires widget env/token |

The normal `ge:skills` flow updates the planner and surface commander used by the add-in. Use the
explicit `--only m365-release-operator` command when you want this setup/deployment skill uploaded as
a separate Gemini Enterprise skill.

## Catalog and tenant deployment

| Need | Command | Notes |
| --- | --- | --- |
| Dry-run unified package catalog upload | `M365_APP_PROFILE=internal-alpha-word-excel scripts/m365-app-catalog.sh upsert --dry-run` | Read-only |
| Upload unified package to catalog | `M365_APP_PROFILE=internal-alpha-word-excel scripts/m365-app-catalog.sh upsert --login` | Uses CLI for Microsoft 365 |
| XML Centralized Deployment dry-run | `scripts/m365-tenant-addin.sh deploy --dry-run` | XML-only; platform support varies |

On Linux, prefer unified package catalog upload for automation. XML Centralized Deployment cmdlets
may require Windows PowerShell depending on Microsoft's module support in the tenant.

