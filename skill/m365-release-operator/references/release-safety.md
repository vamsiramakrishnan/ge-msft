---
title: Release Safety
kind: reference
skill: m365-release-operator
topics: [safety, approval, tenant, entra, catalog, skills, secrets]
load_when: A live command may mutate Entra, Microsoft 365 tenant/catalog state, Gemini Enterprise skills, deployment state, or local credentials.
---

# Release safety

Classify every operation before running it. The harness should display this classification in the UI
and require confirmation for red operations.

## Green: read-only

Examples:

- `bun run setup:doctor`
- `bun run sideload:status`
- `scripts/update-ge-widget-skills.sh --list-only`
- `bun run bootstrap:release:dry-run`

These can run without destructive approval, but may still require login.

## Yellow: local mutation

Examples:

- `bun run setup:package`
- `(cd skill && ./build_zip.sh m365-release-operator)`
- writing `/tmp/ge-widget-token`
- writing `.ge-dev/entra-spa-patch.<profile>.json`

These mutate local files only. They should not touch tenant, catalog, or skill state.

## Red: external mutation

Examples:

- `bun run entra:sync:release`
- `bun run bootstrap:release`
- `scripts/m365-app-catalog.sh upsert --login`
- `scripts/update-ge-widget-skills.sh` without `--list-only`
- `python3 skill/update_skills.py --api-mode widget --replace --yes --live`

These require explicit user approval after a dry-run or target summary. Show the target project,
tenant/profile, manifest/package path, skill labels, and whether old state will be removed.

## Immutable rules

- Never run red operations from ambiguous user text. Ask for confirmation or dry-run first.
- Never delete or replace skills by stale numeric id alone; resolve visible widget skill views first.
- Never register dev tunnel redirects in a release profile.
- Never use `.env` as a secret store.
- Never commit generated token files, HAR files, or raw pasted cURL.
- Never claim deployment success unless the command output proves it.

