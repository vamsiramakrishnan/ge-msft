# checkGrounding — verify claims against facts

Scores how well candidate text is grounded in provided facts; returns per-claim citations. Use to validate agent rewrites before applying a tracked change.

## groundingConfigs.check


Check grounding of an answer candidate against facts.

- **HTTP**: `POST` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/groundingConfigs/{groundingConfigsId}:check`
- **Method id**: `projects.locations.groundingConfigs.check`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

### Request body — `CheckGroundingRequest`

- **userLabels** `object` — The user labels applied to a resource must meet the following requirements: * Each resource can have multiple labels, up to a maximum of 64. * Each label must be a key-value pai…
- **answerCandidate** `string` — Answer candidate to check. It can have a maximum length of 4096 tokens.
- **groundingSpec** `CheckGroundingSpec` — Configuration of the grounding check.
  - **enableClaimLevelScore** `boolean` — The control flag that enables claim-level grounding score in the response.
  - **citationThreshold** `number` — The threshold (in [0,1]) used for determining whether a fact must be cited for a claim in the answer candidate. Choosing a higher threshold will lead to fewer but very strong ci…
- **facts** `array<GroundingFact>` — List of facts for the grounding check. We support up to 200 facts.
  - **attributes** `object` — Attributes associated with the fact. Common attributes include `source` (indicating where the fact was sourced from), `author` (indicating the author of the fact), and so on.
  - **factText** `string` — Text content of the fact. Can be at most 10K characters long.

### Response — `CheckGroundingResponse`

- **supportScore** `number` — The support score for the input answer candidate. Higher the score, higher is the fraction of claims that are supported by the provided facts. This is always set when a response…
- **citedChunks** `array<FactChunk>` — List of facts cited across all claims in the answer candidate. These are derived from the facts supplied in the request.
  - **source** `string` — Source from which this fact chunk was retrieved. If it was retrieved from the GroundingFacts provided in the request then this field will contain the index of the specific fact …
  - **sourceMetadata** `object` — More fine-grained information for the source reference.
  - **uri** `string` — The URI of the source.
  - **index** `integer` — The index of this chunk. Currently, only used for the streaming mode.
  - **domain** `string` — The domain of the source.
  - **title** `string` — The title of the source.
  - **chunkText** `string` — Text content of the fact chunk. Can be at most 10K characters long.
- **citedFacts** `array<CheckGroundingResponseCheckGroundingFactChunk>` — List of facts cited across all claims in the answer candidate. These are derived from the facts supplied in the request.
  - **chunkText** `string` — Text content of the fact chunk. Can be at most 10K characters long.
- **claims** `array<CheckGroundingResponseClaim>` — Claim texts and citation info across all claims in the answer candidate.
  - **startPos** `integer` — Position indicating the start of the claim in the answer candidate, measured in bytes. Note that this is not measured in characters and, therefore, must be rendered in the user …
  - **score** `number` — Confidence score for the claim in the answer candidate, in the range of [0, 1]. This is set only when `CheckGroundingRequest.grounding_spec.enable_claim_level_score` is true.
  - **endPos** `integer` — Position indicating the end of the claim in the answer candidate, exclusive, in bytes. Note that this is not measured in characters and, therefore, must be rendered as such. For…
  - **citationIndices** `array<integer>` — A list of indices (into 'cited_chunks') specifying the citations associated with the claim. For instance [1,3,4] means that cited_chunks[1], cited_chunks[3], cited_chunks[4] are…
  - **claimText** `string` — Text for the claim in the answer candidate. Always provided regardless of whether citations or anti-citations are found.
  - **groundingCheckRequired** `boolean` — Indicates that this claim required grounding check. When the system decided this claim doesn't require attribution/grounding check, this field will be set to false. In that case…

