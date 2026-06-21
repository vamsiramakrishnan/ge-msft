---
name: verify-surface
description: Verify one surface (word, excel, powerpoint, onenote, teams) against its build-plan acceptance criteria and its mockup. Use before marking a surface's exit-gate task done.
---

Given a surface name (word | excel | powerpoint | onenote | teams):

1. Open `docs/mockups/<n>-<surface>.html` and list the signature interactions it demonstrates.
2. Find that surface's tasks and exit gate in `docs/BUILD-PLAN.md`.
3. For each acceptance criterion, run or demonstrate the corresponding behavior in the implemented client (build the package, run its tests, exercise the interaction). Report PASS/FAIL per criterion with the evidence.
4. Confirm the cross-cutting invariants from `docs/02-design.md` hold on this surface: the agent's output lands in the surface's native materials (not just the panel); the unit is shown and composable; changes are provenanced and reversible; identity is scoped to the signed-in user.
5. Summarize: which criteria pass, which fail, and the smallest next change to close each gap. Do not mark the exit gate done if any criterion fails.
