---
title: Example Plan - Excel Analytical
kind: example
skill: m365-command-planner
surface: excel
workflow: single-surface
topics: [excel, analytical, visualization, context]
load_when: A user asks for workbook-scale analysis, reconciliation, charting, or schedule-risk work.
---

# Example plan — Excel analytical task

User:

`/visualize @this find schedule risks across the workbook and create a chart-ready table`

Planner output:

````text
```plan
intent   draft
surface  excel
scope    document
ground   this
context  analytical
context  full-scope
context  upload-preferred
context  code-execution-preferred
step     analyze the workbook for schedule risks and produce a chart-ready summary table
confidence high
```
````

This asks the host/runtime for the right context strategy. It does not invent file ids, upload
commands, Python, or chart commands.
