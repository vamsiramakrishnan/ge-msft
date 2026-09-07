# Task-pane UI

The React view sits over `PanelController`. Host APIs stay in the bridges. Networking, streaming,
conversation state, and approval decisions remain in the controller and runtime.

## Running and inspecting

`bun run dev` starts the standalone preview. Choose **Try interactive demo** to run the real
controller over a scripted session. Switch surfaces and pane dimensions with the preview controls.
Exit demo mode to inspect individual card fixtures. The preview is excluded from the production
Vite entry points. Actual host testing follows the setup guides.

## Component map

| Component | Responsibility |
| --- | --- |
| `App` | Shared dispatch, grounding, planner routing, and busy/approval state wiring |
| `Toolbar` | Context, actions, skills, history, and routing dialog; focus management; Ctrl/⌘K |
| `ContextStrip` | Visible attached-context chips, nearby source discovery, detach and reveal |
| `ContextTray` | Full context list in the toolbar dialog |
| `WorkspaceHome` | Three host-specific starting actions, filtered by capability closure |
| `ActionLibrary` | Search, output filters, and pinned catalog actions |
| `Composer` | Intent, scope, text, keyboard submission, and typed invocation |
| `ComposerSources` | Named source search and request-scoped structured source picks |
| `QuickActionParamForm` | Required values and a preview of the resulting instructions |
| `MessageThread` | Markdown, citations, host locations, completed-artifact insertion, copy and follow-up controls |
| `CommandPlanCard` | Confirmation of the planner's proposed intent and steps |
| `PlanApprovalCard` | Exact effect set, target counts, available before/after values, approval |
| `WriteApprovalCard` / `ShareApprovalCard` | Existing separate approval authorities |
| `ProposalCard` | Proposed/applied changes and provenance; acceptance locks during another turn |
| `RunSteps` | Command execution activity, disclosed on demand |

`SurfaceCommandCenter` and `QuickActionBar` remain exported components with their existing tests;
the main pane uses `WorkspaceHome` and `ActionLibrary` for discovery.

## Data paths

Every action and composer invocation reaches `App.dispatch`. Intent controls and response options
are encoded by `invocationToSeed`; typed source picks are separately resolved by
`invocationToGrounding`. Data-store chips pass resource IDs, never their display labels.

Persistent context chips call controller attach/detach/reveal. Request source chips only modify
the next invocation. Only pinned catalog IDs are stored in local storage. No prompt, document
content, source name, or transcript is persisted by these UI components.

Busy and pending-approval state disables mutation and context controls. The controller also refuses
proposal acceptance while another turn or approval is active. Streaming, failed, and cancelled
artifacts expose no insert control. Output format/style preferences never alter a pasted command
program. Follow-up buttons fill an editable request rather than automatically dispatching.

## Style and accessibility

`styles.css` defines the existing paper/ink/blue palette and base components. `workspace.css` extends
that system for chips, the action library, the start view, and the composer. It is imported after
the base sheet by both production and preview entries. `preview.css` only styles the test harness.

Controls use labels, pressed/expanded states, visible focus, and native input/select elements.
The action dialog traps focus and restores its trigger on close. Hover is supplementary. Context
chips wrap, long labels truncate visually, and the full label remains in the accessible name.
The composer stays reachable while the conversation and decision regions scroll independently.
