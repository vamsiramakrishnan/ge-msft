# Gemini Enterprise on Third‑Party Surfaces

**Surfaces Team — Distribution Plan for M365, Salesforce, and SAP Joule**
*Draft v0.1 · scope: surfacing StreamAssist, no‑code/low‑code agents, and full‑code (ADK) agents into the three ecosystems where enterprise users actually live.*

---

## 1. Thesis: the surface war is already won by protocol, not by plugins

The instinct is to read this as three plugin projects — a Copilot plugin, an AppExchange app, a Joule skill — each a bespoke integration against a foreign extensibility model. That instinct is wrong, and acting on it is how the Surfaces team drowns in three parallel maintenance burdens.

The relevant fact about the market in mid‑2026 is convergence. Microsoft, Salesforce, and SAP have each, independently, stopped exposing their estates to agents as raw request/response APIs and instead opened two sanctioned agentic doorways: **A2A** for agent‑to‑agent delegation and **MCP** for tool/data access. Copilot Studio shipped A2A to GA and added remote MCP. Agentforce 3 made MCP and A2A native and built AgentExchange around them. SAP went furthest — its API Policy now *requires* agentic traffic to route through the Agent Gateway (A2A) or the MCP Gateway in Integration Suite, and forbids agent harnesses from hitting conventional endpoints at all. Three competitors, one substrate.

That substrate is ours. A2A is a Google protocol, donated to the Linux Foundation; ADK has native A2A support and Agent Engine ships A2A endpoints out of the box. So the asymmetry is structural: the thing each host has standardized on as *their* doorway is the thing our agents already speak as a first language.

The plan that follows is therefore not "build three plugins." It is: **expose every Gemini Enterprise agent once, over A2A (primary) and MCP (secondary), then add a thin host‑native wrapper per ecosystem for discovery, packaging, and marketplace distribution.** The protocol layer is built once and amortized across all three surfaces. The per‑host work collapses to registration, identity federation, packaging, and governance — real work, but bounded, and none of it a re‑implementation of the agent.

A blunt way to state the design constraint: *the agent must not know which surface invoked it.* If a Gemini agent contains a single line of M365‑specific or Salesforce‑specific logic, we have built the wrong thing.

---

## 2. What we are surfacing: three modalities, three exposure mechanisms

"Gemini Enterprise agents" is three different artifacts with three different native entry points. Getting the mapping right up front prevents the most common architectural mistake — trying to surface a no‑code agent through the wrong door.

| Modality | Native entry point | A2A‑exposable today? | Mechanism to make it surface‑ready |
|---|---|---|---|
| **Full‑code (ADK) agents** | Agent Engine / Agent Runtime (built‑in A2A endpoints, agent cards + skills, persistent sessions) | **Yes, natively** | Deploy to Agent Engine; it is already an A2A server. Publish the agent card. Zero glue. |
| **No‑code / low‑code (Agent Designer) agents** | `streamAssist` with `agentsSpec` | Not directly | Wrap behind a thin ADK A2A shim whose tool is a `streamAssist` call targeting the agent ID; deploy the shim to Agent Engine. The shim *is* the A2A surface. |
| **StreamAssist (assistant‑level)** | `…/assistants/default_assistant:streamAssist` (streaming, session‑aware, file‑grounded) | Not an agent per se | Surface as a "bring the whole assistant" experience: A2A shim, custom‑engine agent (M365), or invocable service (Salesforce). This is the path when the goal is the full grounded assistant, not one specialized agent. |

Two consequences worth stating plainly:

- The **ADK shim for no‑code agents is the single highest‑leverage piece of engineering in this plan.** Build it once and every Agent Designer agent in every customer tenant becomes A2A‑addressable, and therefore surfaceable into all three hosts, without per‑agent work. This is the multiplier.
- There is a **known dependency risk on `agentsSpec`**: as of early 2026, `streamAssist` was observed ignoring the specified `agentId` (public issue on `google-cloud-python` #16019 and the developer forums). If that is still unresolved, the no‑code → StreamAssist → A2A path is blocked at the source. **Verify the fix before committing the shim to the critical path** (see §8).

---

## 3. The universal substrate (built once, before any host work)

Everything below the host boundary is shared infrastructure. This is Phase 0 and it gates all three ecosystems.

**A2A exposure harness.** A repeatable path that takes any of the three modalities and produces a deployed, discoverable A2A server on Agent Engine with a published agent card (name, skills, auth requirements, streaming semantics). For ADK agents this is near‑zero. For no‑code/StreamAssist it is the shim. Output: a catalog of A2A‑addressable Gemini agents with stable endpoints and cards.

**MCP exposure harness.** For hosts (or customer architectures) that want *tool‑level* rather than *agent‑level* integration, expose Gemini Enterprise tools, retrieval, and grounded knowledge as MCP servers (MCP Toolbox pattern). This is the secondary path and matters most where the host wants to keep its own orchestrator and merely borrow Gemini capabilities as tools.

**Identity & data‑boundary federation.** This is the genuinely hard, genuinely differentiating work, and it is the same problem three times: an end user authenticated in Entra ID / Salesforce Identity / SAP IAS invokes a Gemini agent, and that agent must act with the correct data scope — either *as the end user* (federated identity, so row/document‑level access controls hold) or as a *service principal with an explicit, audited data boundary*. The mechanism is Google Cloud Workforce Identity Federation fronting each host IdP, with a clear contract for which agents may act on‑behalf‑of and which run as service identities. Get this wrong and we either over‑expose data (acting as an over‑privileged service account) or break the experience (no access at all). This workstream is where a serious enterprise buyer will scrutinize us hardest.

**Two‑plane separation.** Keep the *control plane* (agent registration, marketplace listing, admin governance, package lifecycle) cleanly separate from the *data plane* (A2A/MCP runtime invocation, streaming, identity propagation, session/state handoff). The hosts model this distinction internally; mirroring it keeps our integrations debuggable and lets governance evolve independently of runtime.

---

## 4. Ecosystem 1 — Microsoft 365

**Host model.** M365 Copilot offers two build modes: *declarative agents* (customize instructions/knowledge/actions on top of Copilot's own orchestrator and model — manifest schema v1.7, built via Agent Builder no‑code or the Agents Toolkit pro‑code) and *custom‑engine agents* (bring your own orchestrator and model via the M365 Agents SDK). A2A is GA in Copilot Studio and lets agents delegate to first/second/third‑party agents over the open protocol; remote MCP servers and MCP Apps (interactive UI returned into Copilot Chat) are supported. Registration and admin control run through the Agent Registration API and the Package Management API. Distribution is via the Agent Store / Microsoft commercial marketplace and Teams.

**Integration paths, ranked.**

1. **A2A delegation (primary).** Register the Gemini A2A agent as a third‑party A2A agent in Copilot Studio; Copilot's orchestrator delegates tasks to it. Works for full‑code and (via the shim) no‑code agents. This is the path that preserves Gemini *reasoning* — the agent thinks with Gemini, not with Copilot's model.
2. **Custom‑engine agent over StreamAssist.** Build a custom‑engine agent on the M365 Agents SDK that proxies StreamAssist, surfacing the *full grounded Gemini assistant* inside Teams, Outlook, and Copilot Chat with our orchestrator and model. Use this when the deliverable is the whole assistant experience rather than one specialist.
3. **Declarative agent + API plugin / remote MCP.** Expose Gemini tools/knowledge as actions or an MCP server consumed by a declarative agent. Note the limitation honestly: a declarative agent runs on *Copilot's* model, so this path surfaces Gemini *capabilities*, not Gemini *intelligence*. Fine for grounding and tool‑calling; not a model‑differentiation play.
4. **MCP App.** A Gemini‑backed MCP server that returns interactive UI (approval forms, dashboards) into Copilot Chat. Use where the surface needs rich in‑chat UX, not just text.

**Identity.** Entra ID OAuth, on‑behalf‑of flow → Workforce Identity Federation.

**Surfaces reached.** Copilot Chat, Teams, Outlook, the M365 app shell.

**Risks specific to M365.** (a) The declarative path gives no model differentiation — resist customer/field pressure to default to it. (b) Copilot's own guidance flags prompt‑injection from untrusted content (emails, tickets) reaching agent tools; our exposed tools must assume hostile input. (c) Data residency: M365's regional commitments vs Gemini's regional endpoints — map per JAPAC geography, with explicit attention to India and Australia residency expectations.

**Distribution & governance.** Agent Store / commercial marketplace listing; tenant admins gate availability via the Package Management API (block/unblock, metadata, ownership). Plan for admin‑side controls as a first‑class requirement, not an afterthought.

---

## 5. Ecosystem 2 — Salesforce

**Host model.** Agentforce 3 made MCP and A2A native (Salesforce is a founding member of both). The Atlas Reasoning Engine 3.0 supports calling external MCP servers as action sources and orchestrates multi‑agent delegation (orchestrator + specialist agents, context handoff — GA in Summer '26). AgentExchange is the agentic marketplace: pre‑built agents, skills, and 50+ MCP servers from 200+ partners. Headless 360 (TDX 2026) exposes Salesforce capabilities as MCP tools to external agents. MuleSoft offers zero‑code API‑to‑MCP conversion; Heroku can host custom MCP servers. One material constraint: each Agentforce agent supports roughly **20 simultaneous MCP tools**, which forces curated, focused toolsets.

**Integration paths, ranked.**

1. **A2A specialist (primary).** Expose the Gemini agent as an A2A agent with an agent card; the Agentforce orchestrator (Atlas 3.0) delegates to it as a specialist. This is the path that uses Gemini's *reasoning* as a distinct agent in the team, rather than overloading a single Salesforce agent's context. Best for full‑code and wrapped no‑code agents.
2. **MCP server / actions.** Register the Gemini‑backed MCP server in Setup with allowlists, add it as MCP actions to an Agentforce agent. Respect the ~20‑tool ceiling — curate tightly. Best for surfacing Gemini *tools/retrieval*.
3. **Apex / External Services over StreamAssist.** Invocable Apex or an External Service calling StreamAssist. Maximum control, works in any org including those not yet on Agentforce 3 — the dependable fallback and the right choice for classic Lightning embedding.
4. **AgentExchange / AppExchange managed package + LWC.** Package a Gemini‑powered experience with Lightning Web Components for richer in‑Salesforce‑UI embedding plus a marketplace distribution channel.

**Identity.** Salesforce OAuth / External Client Apps / Named Credentials → Workforce Identity Federation.

**Surfaces reached.** Agentforce across Sales and Service Cloud, Lightning experiences, and Slack (Salesforce‑owned, a meaningful bonus surface).

**Risks specific to Salesforce.** (a) The 20‑tool/agent MCP cap — design for curation, not breadth. (b) Agentforce 3's trust layer means identity and granular access control are enforced and scrutinized; our federation must satisfy it. (c) AgentExchange security review (tool poisoning, supply‑chain) is a gate on marketplace listing — budget for it.

**Distribution & governance.** AgentExchange for agentic listings; AppExchange managed package for packaged apps. Both carry security review.

---

## 6. Ecosystem 3 — SAP Joule

**Host model.** Joule Studio (GA, with 2.0 unveiled at Sapphire 2026) builds no‑code and pro‑code agents on BTP with a fully managed runtime, CLI, and VS Code/Cursor tooling. SAP recommends exactly two governed agentic pathways: the **Agent Gateway (A2A)** for multi‑agent collaboration where an external/third‑party agent delegates to or receives results from SAP‑managed agents, and the **MCP Gateway in Integration Suite** for governed exposure/consumption of SAP and non‑SAP APIs as MCP tools. A2A is **bidirectional** — SAP's own materials call out that third‑party agents built on Google's framework, Copilot Studio, or any A2A‑compatible platform can natively discover and invoke Joule agents, and Joule (as an A2A client) can invoke external agents. SAP publishes reference implementations for A2A‑compliant pro‑code agents with Joule integration on BTP.

**The defining constraint — and why it favors us.** SAP's API Policy mandates that agentic traffic route *only* through these two gateways; it explicitly does not permit agent harnesses to hit conventional OData/request‑response endpoints. For most vendors this is friction. For an A2A‑native agent platform it is a moat in our favor: the only sanctioned way in is the door we already speak through, and SAP's reference architecture names exactly our approach. Lean into this.

**Integration paths, ranked.**

1. **A2A via Agent Gateway (primary, and SAP‑sanctioned).** Expose the Gemini agent as an A2A server; register through the Agent Gateway; Joule discovers and delegates to it. This matches SAP's published reference architecture for external pro‑code agents. Best for full‑code and wrapped no‑code agents.
2. **MCP via MCP Gateway (Integration Suite).** Expose Gemini tools as an MCP server, governed through the MCP Gateway, consumed by Joule agents. Best for tool‑level integration.
3. **Joule Studio agent‑as‑a‑tool / pro‑code extensibility on BTP.** Build a Joule Studio agent that wraps a Gemini agent via the gateways above, using SAP's upgrade‑safe (clean‑core) extensibility so customer landscapes stay upgradeable.

**Identity.** SAP IAS / XSUAA with principal propagation; Agent Gateway handles A2A identity → Workforce Identity Federation.

**Surfaces reached.** Joule across S/4HANA, SuccessFactors, Ariba, and the broader SAP estate — i.e., the systems of record where the most consequential enterprise actions happen.

**Risks specific to SAP.** (a) API Policy compliance is non‑negotiable — there is no "just call the API" escape hatch; all paths go through the gateways. (b) Clean‑core / upgrade‑safe extensibility constraints on anything touching Joule agents. (c) BTP deployment footprint and data residency mapping. (d) Commercial: SAP is funding this ecosystem (€100M partner fund, free design‑time access, Hack2Build, CodeJam Hyderabad on 23 June 2026) — a partner motion here is cheap leverage, and the Google–SAP relationship should be used to formalize it.

**Distribution & governance.** SAP Store / BTP; pursue the partner fund and Hack2Build track as a low‑cost accelerant.

---

## 7. Cross‑cutting workstreams

These run across all three ecosystems; most are built once.

- **WS1 — Protocol & SDK foundation.** The A2A exposure harness (incl. the no‑code/StreamAssist ADK shim), the MCP exposure harness, the agent‑card catalog, and reusable server templates. *Built once; serves all hosts.* Highest leverage in the plan.
- **WS2 — Identity & data‑boundary federation.** Workforce Identity Federation fronting Entra / Salesforce Identity / SAP IAS; the on‑behalf‑of‑vs‑service‑principal contract; audited data boundaries. *The hardest and most differentiating workstream.*
- **WS3 — Host‑native packaging & marketplace.** Declarative‑agent/custom‑engine packaging (M365), AgentExchange/AppExchange package + LWC (Salesforce), Joule Studio agent + BTP deploy (SAP); plus admin/governance surfaces per host.
- **WS4 — Security & governance.** Prompt‑injection defense (Model Armor at the boundary), MCP tool‑poisoning and supply‑chain controls, allowlists, end‑to‑end audit logging, RAI/compliance posture, and a per‑region data‑residency map (explicit India and Australia treatment for JAPAC).
- **WS5 — Observability & evaluation.** Cross‑surface tracing (one trace from host invocation through Gemini agent and back), a per‑surface eval harness, and latency budgets honoring streaming/SSE first‑token targets — the experience is judged on responsiveness, not just correctness.
- **WS6 — Reference implementations & DX.** "Surface‑in‑a‑box" templates, sample agents, docs, and codelabs — mirroring how M365, Salesforce, and SAP themselves ship reference implementations. This is how the field and partners self‑serve instead of escalating.

---

## 8. Sequencing

**Phase 0 — Foundation (gate for everything).** Build WS1 (A2A + MCP harnesses, the no‑code shim, the agent‑card catalog) and stand up WS2 identity federation in skeleton form. **Hard dependency check first:** confirm `streamAssist` `agentsSpec` correctly targets agent IDs (issue #16019); confirm Agent Engine A2A endpoints are GA in the target regions; confirm agents are on the current Gemini Enterprise Agent Platform SDK ahead of the Vertex‑AI SDK deprecation (deprecated modules stop updating 24 June 2026). Ship one lighthouse agent into one ecosystem over pure A2A to prove the spine end‑to‑end.

**Phase 1 — One lighthouse per ecosystem.** Replicate the A2A path across all three hosts with a single, real customer agent each. Sequence by customer pull and strategic value rather than by host readiness — all three are protocol‑ready today, with SAP and Salesforce arguably the most A2A‑mature. Harden identity federation against real tenant access controls. Add MCP tool exposure where the host wants tools rather than an agent.

**Phase 2 — Packaging & marketplace.** Host‑native wrappers (WS3), marketplace listings and their security reviews, and admin/governance surfaces (WS4). This is where "works in a demo" becomes "an admin can safely turn it on for 50,000 seats."

**Phase 3 — Scale (bellwether‑and‑replicate).** Generalize from the lighthouses, turn on observability/eval (WS5), publish reference implementations and DX (WS6), and drive toward GA across surfaces. Target metric: time‑to‑surface a new agent measured in hours via the shim, not weeks of bespoke integration.

---

## 9. Risks & open dependencies

- **`agentsSpec` ignored in StreamAssist (early‑2026 bug).** If unfixed, the no‑code → StreamAssist → A2A spine is blocked. *Verify before committing the shim to the critical path.* Mitigation: front the agent via an ADK wrapper that invokes the agent directly rather than through assistant‑level routing.
- **Declarative agents run the host model (M365).** No Gemini model differentiation on that path. Default to A2A or custom‑engine for genuine Gemini reasoning; reserve declarative for grounding/tools.
- **Agentforce ~20‑tool MCP cap.** Curate aggressively; do not design for tool breadth.
- **SAP API Policy.** No raw‑endpoint path exists; all integration goes through Agent Gateway (A2A) or MCP Gateway. (Net positive for us, but it removes the usual fallback.)
- **Identity federation across three IdPs.** The hardest engineering and the sharpest buyer scrutiny. Under‑invest here and the whole motion stalls at security review.
- **Data residency (JAPAC).** Host residency commitments vs Gemini regional endpoints; India and Australia need explicit mapping given customer sensitivity. Tie to existing residency competitive work.
- **Protocol drift.** A2A and MCP are LF‑governed and actively evolving. Pin versions, track spec changes, and run the A2A Inspector / compatibility kit in CI.
- **The "seam" / orchestrator handoff problem.** When a host orchestrator delegates to a Gemini agent, context and session state must transfer cleanly across two independent stacks. Define an explicit handoff contract (what context crosses, who owns memory, how interruptions propagate). All three hosts have this seam; it is the most common source of multi‑agent flakiness.
- **Marketplace security reviews** (AgentExchange, AppSource) are gates with lead time — start them early, not at launch.

---

## 10. How we'll know it's working

- **Coverage:** number of Gemini agents live per ecosystem, and **time‑to‑surface a new agent** (target: hours via the shared harness).
- **Reuse ratio:** share of host integrations served by the common A2A/MCP foundation vs bespoke code. *This is the number that proves the thesis;* if it trends down, we are quietly rebuilding three plugins after all.
- **Usage:** invocations and MAU per surface; task completion / deflection on real workflows.
- **Identity integrity:** share of sessions correctly acting as the end user (with access controls enforced) vs service principal.
- **Distribution:** marketplace listings and installs across Agent Store, AgentExchange, and SAP Store.
- **Latency:** streaming first‑token and end‑to‑end times per surface, against budget.

---

## 11. The one‑sentence version

Build the A2A/MCP exposure spine once — with the no‑code ADK shim and cross‑host identity federation as its load‑bearing pieces — and let three competitors' own sanctioned agentic gateways carry Gemini Enterprise agents into Copilot, Agentforce, and Joule; everything host‑specific is distribution, not engineering.

---

### Sources (product‑state references, June 2026)

Gemini Enterprise / ADK / A2A:
- StreamAssist API — Google Cloud docs: https://docs.cloud.google.com/gemini/enterprise/docs/get-answers-from-streamassist
- `agentsSpec` issue — https://github.com/googleapis/google-cloud-python/issues/16019
- ADK on Gemini Enterprise Agent Platform — https://docs.cloud.google.com/gemini-enterprise-agent-platform/build/adk
- A2A native in ADK + Agent Engine — https://cloud.google.com/blog/products/ai-machine-learning/agent2agent-protocol-is-getting-an-upgrade
- Expose ADK agent as A2A server on Agent Runtime (codelab) — https://codelabs.developers.google.com/adk-a2a-agent-runtime
- Gemini Live API (voice/video surfaces) — https://docs.cloud.google.com/gemini-enterprise-agent-platform/models/live-api

Microsoft 365:
- Agents overview (declarative vs custom‑engine) — https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/agents-overview
- Copilot extensibility overview (A2A / MCP) — https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/overview
- Copilot Studio multi‑agent / A2A GA — https://www.microsoft.com/en-us/microsoft-copilot/blog/copilot-studio/new-and-improved-multi-agent-orchestration-connected-experiences-and-faster-prompt-iteration/
- What's new (Agent Registration API, Package Management API) — https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/whats-new

Salesforce:
- Connected agents / AgentExchange (MCP + A2A) — https://www.salesforce.com/blog/connected-agents-agentexchange/
- Choosing API vs MCP vs A2A — https://www.salesforce.com/blog/how-to-choose-integration-pattern-for-agentforce/
- Agentforce MCP guide (Headless 360, tool cap, registration) — https://vantagepoint.io/blog/sf/agentforce-mcp-connect-external-systems-guide

SAP Joule:
- A2A & MCP for interoperability / Agent Gateway + MCP Gateway — https://architecture.learning.sap.com/docs/ref-arch/ca1d2a3e/1
- Integrating AI agents with Joule (A2A reference impls) — https://architecture.learning.sap.com/docs/ref-arch/ca1d2a3e/4
- Joule Studio agent builder GA (MCP, agent‑as‑tool, A2A roadmap) — https://community.sap.com/t5/technology-blog-posts-by-sap/build-business-grounded-ai-agents-faster/ba-p/14258166
- SAP API Policy / agentic gateway routing — https://www.theregister.com/saas/2026/05/19/saps-ai-strategy-come-for-the-openness-stay-because-you-have-to/5241109
