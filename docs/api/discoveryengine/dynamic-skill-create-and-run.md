# Dynamically creating a skill and running it — what works, what doesn't (live-verified)

Goal: create a skill agent at runtime via the public API and immediately invoke it. **Verdict (saib,
2026-07-08): you can CREATE dynamically, but a freshly-created skill does NOT become invocable/routable
— the assistant tags it a "dynamic skill" and declines. Established skills (planner/commander) route
fine.** So "create and immediately run" is **not reliable** on this tenant via the public path.

## The workflow (as proposed) + what each step actually returns

All against `…/engines/ge-msft-plugin-test_1782382759735/assistants/default_assistant`, WIF token,
`X-Goog-User-Project: saib-ai-playground`.

### Step 1 — CreateAgent (dynamic) — ✅ works
```bash
AID="dyn-skill-$RANDOM"
curl -s -X POST "$BASE/agents?agentId=$AID" -H "Authorization: Bearer $TOK" \
  -H "X-Goog-User-Project: saib-ai-playground" -H "Content-Type: application/json" -d '{
    "displayName":"'"$AID"'","description":"…","state":"PRIVATE",
    "skillAgentDefinition":{"instruction":"…"}}'
```
→ **HTTP 200**. The agent then appears in `:listAvailableAgentViews` and `GetAgent` returns it.

### Step 1b — bundle upload (optional) — ✅ works, and *validates*
`POST …/agents/$AID/files:upload?upload_protocol=raw` (`Content-Type: application/zip`) → **200**;
`SKILL.md` (with full frontmatter `name/description/license/allowed-tools`) → `instruction`, rest →
`subfiles`. Confirmed via GetAgent. So the format is accepted.

### Step 2 — invoke — ❌ does NOT route a freshly-created skill
- **Method A — `mention://?uri=$AID` in `query.text`:** returns 200, but `invokedSkills=(none)` and the
  answer is the **default assistant** saying *"👋 It looks like you've mentioned the dynamic tool/skill
  **dyn-…***" — it sees the mention and **declines to execute it**. Tested bare-instruction and
  fully-bundled-with-frontmatter, and across 4 attempts over 24s — never routed.
- **Method B — `agentSelectionConfig.enabledAgentIds`:** **HTTP 400** `Invalid JSON payload … Unknown
  name "agentSelectionConfig"`. **This field does not exist** in public v1alpha. (Same as `agentsSpec`/
  `agentsConfig`: not on the reachable surface.)

### Step 3 — DeleteAgent — ✅ works
`DELETE …/agents/$AID` → **200**. (Deletes the slug id; that slug is then soft-tombstoned — don't reuse.)

## The isolation that proves it's not the obvious causes

| Test | Result |
|---|---|
| Dynamic skill present in `:listAvailableAgentViews`? | **yes** (registered + available) |
| CONTROL: mention **commander** (numeric id `3708…`), actionable prompt | **routes** → `invokedSkills=[m365-surface-commander]` |
| Mention **planner** (SLUG id `m365-command-planner`) | **routes** → `invokedSkills=[m365-command-planner]` |
| Mention **fresh dynamic** skill (bare instruction) | **no route** — "dynamic skill", declined |
| Mention **fresh dynamic** skill (full frontmatter bundle, upload 200) | **no route** — "dynamic skill", declined |

So it is **not**: id format (slug routes — planner), frontmatter/agents-spec validity (bundled valid skill
still declined), availability (it's listed), or short propagation (24s+). The only remaining difference is
that the **established** skills have crossed some **router-eligibility / enablement / indexing** state that
a freshly public-created skill has not.

## Open question for the DE team

What makes a skill **mention-routable** by the assistant, beyond existing + being in
`listAvailableAgentViews`? Candidates (unconfirmed): a long router-indexing latency (minutes–hours), an
explicit enable/deploy/share/review step the widget performs but public `CreateAgent` does not, or a
widget-origin requirement. The assistant's own copy ("you've mentioned the dynamic tool/skill …")
suggests a first-class distinction between "dynamic/unenabled" and routable skills.

## Practical guidance (what to actually do)

- **Do not rely on create-then-immediately-run.** For the add-in, **pre-provision** the skills (our
  boot-time `ensureSkillAgent` warm-up) — they route once established, as planner/commander do. Give a
  newly-provisioned skill time / whatever enablement the established ones had before depending on routing.
- **Invocation lever that works:** `mention://?uri=<id>` for an **established** skill (numeric or slug id).
- **Not real / not working:** `agentSelectionConfig`, `agentsSpec`, `agentsConfig` (all 400/500/ignored).
- Lifecycle (create/get/update/delete/upload) is fully public + WIF; only *routability of fresh skills* is
  the gap. See [[ge-agent-invocation-live-findings]] and listing-agents-and-skills.md.
