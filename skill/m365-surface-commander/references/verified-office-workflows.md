---
title: Verified Office Workflows
kind: reference
skill: m365-surface-commander
topics: [freshness, bounded-context, managed-refresh, native-verification, recovery]
load_when: A task needs reusable Office workflows, managed output refresh, native postconditions or interrupted-write recovery.
---

# Verified Office workflow rules

- Available writes are the current effective manifest: host requirements, item mode and release
  profile can remove a command. Never try a hidden operation through a generic invocation.
- A paged evidence query can return no matches and still have `nextCursor`. Advance until
  `complete` or an explicit budget/coverage limit. Unit/text-window reads verify their stored page.
- Excel cell writes with unfamiliar or array formula output shapes require staging. Do not replace
  an existing unknown spill formula by writing only its parent cell. Use bounded captured data and
  deterministic analysis when appropriate. A computed grid is not an Excel PivotTable.
- Word current-revision readback excludes tracked deletions. Preserve the user's tracking mode.
  Report extraction limitations: body paragraph coverage does not establish headers/footers,
  fields, table structure or whole-file coverage.
- Managed Word content-control/PPT shape refresh uses a native capture and exact object identity.
  Never invent `expectedDocumentId` or `expectedTextHash`. The task-pane/SDK workflow supplies
  these restrictions. A manual edit is a conflict, not permission to overwrite it.
- An Outlook review form being opened is not a sent message or proof of a saved draft. Say the
  actual observed state. Compose-only edits operate on the open draft, not another window.
- Preserve unknown outcomes. A persisted journal claim is not an approval; do not resend an
  uncertain effect or bypass inspection. Continuation requires fresh context and normal approval.
- Deck charts are editable native artifact objects. Theme/geometry checks do not establish font
  overflow, native master fidelity or successful insertion. Only the host's receipt and applicable
  postcondition can establish completion.
