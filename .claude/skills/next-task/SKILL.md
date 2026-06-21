---
name: next-task
description: Pick the next unchecked task from docs/BUILD-PLAN.md, plan it, implement it against the contracts and conventions, verify its acceptance criteria, and check it off. Use this to drive the build.
---

Drive the build forward by one task.

1. Read `docs/BUILD-PLAN.md` and find the first task whose checkbox is `[ ]` (skip `[x]` and `[~]`).
2. If this task is the first in a new phase, enter plan mode (`/plan`) and scope the phase before writing any code. Use the Explore subagent to scan relevant existing code so the main context stays clean.
3. Read the parts of `docs/CONTRACTS.md` and `docs/CONVENTIONS.md` relevant to this task. If the task names a surface, open the matching `docs/mockups/*.html` to confirm the intended interaction before building.
4. Implement the task. Touch only the package(s) the task names. Keep `packages/web-shell` surface-agnostic — surface specifics go in `bridge-*` or `teams`.
5. Verify: run `npm run typecheck`, `npm run test`, and `npm run lint`. Confirm the task's acceptance criterion (AC) is actually met — run or demonstrate it, don't assume.
6. If the task touched authentication, credentials, identity federation, guardrails, or provenance, invoke the `security-reviewer` subagent and resolve any findings before continuing.
7. If a required external resource (GCP/Entra/Gemini) isn't available, implement against the contract with a clearly-labelled mock, set the checkbox to `[~]`, note the blocking dependency, and stop. Otherwise change the checkbox to `[x]`.
8. Commit with a message referencing the task id (e.g. `feat(gateway): 0.6 StreamAssist SSE relay`). Summarize what changed and what the next task is.

Do exactly one task. Don't batch ahead.
