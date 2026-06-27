# A2UI — agent-authored interactive UI, mapped to host actuations

A2UI ("agent-to-UI") is Google's open protocol (co-developed with the Flutter team and the
Gemini Enterprise team) that lets an agent return a **declarative component tree** instead of
text/HTML. Gemini Enterprise renders it natively; **our add-in renders it itself** (we're inside
Word/Excel/Outlook, not GE chat) and maps its actions onto our capability registry.

Sources: [A2UI spec/quickstart](https://a2ui.org/quickstart/) ·
[A2UI + Gemini Enterprise integration](https://cloud.google.com/blog/topics/developers-practitioners/guide-to-gemini-enterprise-and-a2ui-integration) ·
[Register & manage A2UI agents](https://docs.cloud.google.com/gemini/enterprise/docs/a2ui-agents/register-and-manage-an-a2ui-agent).

## What it is

- The agent emits a **component tree** (`Card`, `Text`, `Button`, `ChoicePicker`, `Image`, …) plus
  a **separate data model** holding the values components display.
- It's a **flat list of small JSON messages**, streaming-friendly — the LLM emits incrementally and
  the client paints as they arrive.
- It's **pure data, catalog-validated**: the client only renders components from a pre-approved
  catalog, so a remote agent cannot inject code or exfiltrate credentials. No `eval`, ever.
- Framework-agnostic: the same payload renders in Lit/Angular/Flutter/native — or, for us, Fluent/
  Office UI in a task pane.

## Wire format (v0.9.1)

Components (structure):
```json
{
  "version": "v0.9.1",
  "updateComponents": {
    "surfaceId": "main",
    "components": [
      { "id": "why", "component": "Text", "value": { "path": "/finding/why" } },
      {
        "id": "rewrite", "component": "ChoicePicker",
        "options": { "path": "/finding/rewrites" }, "selection": { "path": "/finding/chosen" }
      },
      { "id": "apply", "component": "Button", "child": "apply-label", "variant": "primary",
        "action": { "event": { "name": "host:tracked-change" } } },
      { "id": "explain", "component": "Button", "child": "explain-label",
        "action": { "event": { "name": "explain" } } }
    ]
  }
}
```

Data model (values, sent independently so later turns can update data without re-sending the tree):
```json
{
  "version": "v0.9.1",
  "updateDataModel": {
    "surfaceId": "main",
    "path": "/finding",
    "value": { "why": "Below the FSI availability floor.",
               "rewrites": ["99.9% of the time", "99.95% of the time"], "chosen": 0 }
  }
}
```

A `Button` declares its action as `"action": { "event": { "name": "<ACTION_ID>" } }`. On
interaction, the renderer sends back the **action name + the bound data-model values** as the
payload.

## Transport through Gemini Enterprise

- Agent → GE: A2UI rides inside the **A2A protocol** as `DataPart` objects with MIME
  `application/json+a2ui`. GE sends the agent its **catalog** of approved components; the agent
  emits A2UI **or** falls back to text (both can coexist in one response).
- User interaction → agent: serialized as JSON and returned on the **next `streamAssist` turn** via
  `query.parts[].uiJsonPayload` (confirmed in the `v1alpha` schema: *"As of Q1 2026,
  `ui_json_payload` is only supported for A2UI messages."*).
- Response-side carrier in `streamAssist`: the agent's A2UI is delivered in the assistant content of
  the streamed reply (an `AssistantContent` part carrying the `application/json+a2ui` data). Treat
  the exact field defensively — detect the `application/json+a2ui` MIME / an `updateComponents`/
  `updateDataModel` envelope rather than assuming a fixed property.

## How WE use it (add-in renderer + action router)

Because the add-in lives in an Office host, it runs its own A2UI runtime:

1. **Render** only allowlisted catalog components → Fluent/Office UI. Strings are data, never markup.
2. **Route** a fired `action.event.name`:
   - **`host:*`** → execute locally as an **actuation**. Look the name up in the surface's
     `CapabilityManifest`; build an `ActuationRequest` from the bound data-model values; `actuate()`.
     Example: `host:tracked-change` + `{ chosen rewrite, target }` → a tracked change in Word, or a
     cell write in Excel — *same agent, different host*. Provenance is stamped from the turn's
     `provenance` event (agent id, sources, session).
   - **anything else** → round-trip to the agent: serialize `{ event, dataModel }` into
     `query.parts[].uiJsonPayload` and send the next `streamAssist` turn.
3. **Guardrails**: only execute `host:*` actions that exist in the manifest (an agent cannot invent
   a destructive host action); warn before non-reversible ones; cap payload sizes.

### The action-manifest convention

Each custom agent declares which A2UI event names are host actuations vs agent-internal, by
namespacing: `host:<ActuationKind>` is locally executable and its payload shape must satisfy
`ActuationParams`; un-namespaced events round-trip. This keeps the agent **app-independent** — it
authors intent-level UI; the add-in binds each action to the right host capability per surface.

## Why this matters

A2UI is the missing third leg: connectors/agents are **request fields** (see
`context-mechanisms.md`, `agent-invocation.md`), content quality is **bidirectional metadata**, and
A2UI lets a custom agent *propose* interactive, reviewable UI while our capability registry
*executes* it — reversibly and provenanced — in whatever Microsoft app the user is in.
