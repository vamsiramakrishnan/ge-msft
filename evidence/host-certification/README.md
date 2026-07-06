# Host-certification evidence

This directory holds the **live-host certification reports** that the release gate reads. The
protocol, the checklist, and the exact JSON shape are defined in
`docs/release/HOST-CERTIFICATION.md`; the gate that consumes them is `certificationOutcomes()` in
`tools/release/release-check.mjs`.

## How it works

- One JSON file per certified surface (× platform run, if you split them). The gate scans every
  `*.json` under this directory and matches on the file's `surface` and `profile` fields — the
  filename is a convention for humans, not parsed by the gate.
- Naming convention: `<profile>.<surface>.v<manifestVersion>.<platform>.json`, e.g.
  `internal-alpha-word-excel.word.v0.1.0.desktop-windows.json`
  `production.excel.v1.2.0.web.json`
- Each report is pinned to the exact release by `commitSha`, `packageHash`, and
  `manifestVersion`. A report from an older commit or package simply stops matching — leave old
  reports in place as history; they do not interfere.

## The gate stays BLOCKED until real evidence exists

`npm run release:check` reports `live-host certification word` / `... excel` as **BLOCKED** while
no matching report exists for the release under check. That is the correct, honest state until a
human has actually run the protocol on real Word/Excel hosts with the exact built package.

**Do not fabricate evidence files.** Do not write reports for runs that did not happen, copy an
old report and update its hashes, or generate reports from CI. A report in this directory is an
attestation by the named certifier that every checklist item passed on a real host against the
exact bytes identified by `commitSha` + `packageHash`. Agents working in this repo must never
create `*.json` files here; only a human who performed the certification run does.
