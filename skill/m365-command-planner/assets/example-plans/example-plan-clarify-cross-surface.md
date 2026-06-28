---
title: Example Plan - Clarify Cross-Surface
kind: example
skill: m365-command-planner
workflow: cross-surface
topics: [clarification, handoff, target-surface]
load_when: A cross-product request lacks target surface, scope, or approval boundary.
---

# Example plan — clarify before cross-surface work

User:

`/draft turn this into a client update and send it around`

Planner output:

````text
```plan
intent   draft
surface  excel
scope    selection
workflow cross-surface
source   excel selection
context  inline-preferred
clarify  Which target should receive the client update: PowerPoint deck, Word memo, Outlook draft, or Teams post?
clarify  "send it around" would be externally visible; should I only prepare a draft for review?
confidence low
```
````

Why this is correct:

- The user did not name the target surface.
- The request implies an external send, so dispatch must stop at clarification.
- No `target`, `phase`, or `handoff` is invented before the user answers.
