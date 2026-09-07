---
title: Example Plan - Excel To PowerPoint
kind: example
skill: m365-command-planner
surface: excel
workflow: cross-surface
topics: [excel, powerpoint, handoff, deck-generation]
load_when: A user asks to create a PowerPoint deck from an Excel workbook or range.
---

# Example plan — Excel workbook to PowerPoint deck

User:

`/draft make an executive PowerPoint deck from this workbook, using top risks, timeline, and status`

Planner output:

````text
```plan
intent   draft
surface  excel
scope    document
context  analytical
context  full-scope
context  upload-preferred
context  code-execution-preferred
step     Analyze the workbook in Excel, extract top risks, timeline, and status, and prepare a slide-ready handoff packet
step     After the user opens PowerPoint and reviews the handoff, create an executive deck from the approved handoff packet
step     Include in the reviewable handoff: chart-ready summary table, slide outline, source workbook refs, exclusions, and provenance
step     prepare the Excel analysis handoff; do not mutate PowerPoint from Excel
exclude  do not overwrite source workbook cells
confidence high
```
````

Why this is correct:

- `surface excel` says where the workflow starts.
- Ordered steps make the handoff and separate app review explicit.
- The Excel phase prepares a packet; it does not create PowerPoint slides.
- PowerPoint receives its own later commander run and approval.
