---
title: Example Plan - Word Rewrite
kind: example
skill: m365-command-planner
surface: word
workflow: single-surface
topics: [rewrite, tracked-change, selection]
load_when: A user asks for a single-surface Word rewrite or tracked change.
---

# Example plan — Word rewrite

User:

`/rewrite the SLA availability figure to our 99.9% standard, as a tracked change`

Planner output:

````text
```plan
intent   rewrite
surface  word
scope    selection
context  inline-preferred
step     rewrite the SLA availability figure to 99.9% as a tracked change
confidence high
```
````

The planner does not guess the current SLA text. The commander must read/search the live document
and anchor the tracked change on exact text.
