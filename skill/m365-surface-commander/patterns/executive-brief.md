# Pattern: executive brief (cross-surface handoff)

A reasoning template, not a command. Read it for shape, then write the turn's algebra.

**Intent:** "analyze this data and turn it into an exec slide / brief."

This pattern **crosses surfaces** (Excel → PowerPoint), so it is deliberately **two phases**, not one
program — the artifact/surface change is a break boundary (ADR-0008 §3). Each surface is its own
add-in instance with its own approval; you cannot pretend one transaction spans both.

**Phase 1 — analyze in Excel (OBSERVE → DERIVE → EFFECT)**

```
let $top = read Sales!A1:G5000 | filter … | sort <measure> desc | head 5 | select <cols>
spill Brief!A1 = ($top)
table Brief!A1:C6 headers
chart bar Brief!A1:C6 title="…"
done
```

Materialize the analysis where it lives. The derived numbers are now concrete and reviewed.

**Phase 2 — compose in PowerPoint (a fresh turn / signed handoff)**

```
slide "Q3 in one slide" "Top 5 by revenue: …" "Key driver: …" "Risk: …"
# or a richer fragment via the specialized surface, when advertised:
/insert-image base64=<chart png> alt="Top 5 by revenue"
```

Carry the **derived facts** (the numbers you computed in Phase 1), not a live link — the two surfaces
don't share state.

**Failure rule:** never present Phase-2 slide numbers you didn't compute and review in Phase 1.

**Anti-patterns**

- trying to drive Excel analysis and PowerPoint authoring in one `cmd` program (one fake transaction
  across two add-ins);
- charting raw source for a "top 5" instead of the derived range;
- restating numbers in the slide that were never materialized/reviewed.
