
# StreamAssist — the streaming grounded assistant

The primary chat/grounding entry point. The assistant (agents, Model Armor, grounding data stores) is configured **in the Gemini Enterprise engine** — the client only sends a query + optional session + tool/grounding scope. Note: in `v1alpha` there is **no** `agentsSpec` field on the request (the known agent-id bug is avoided by configuring routing on the assistant). The response streams as chunked JSON objects; accumulate `answer.replies[].groundedContent.content.text` for tokens and read `...textGroundingMetadata.references[]` for citations.

## streamAssist


Assist a query in streaming fashion.

- **HTTP**: `POST` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/collections/{collectionsId}/engines/{enginesId}/assistants/{assistantsId}:streamAssist`
- **Method id**: `projects.locations.collections.engines.assistants.streamAssist`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.assist.readwrite, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

### Request body — `StreamAssistRequest`

- **session** `string` — Optional. The session to use for the request. If specified, the assistant has access to the session history, and the query and the answer are stored there. If `-` is specified a…
- **actionSpec** `StreamAssistRequestActionSpec` — Optional. Specification of actions for the request.
  - **actionDisabled** `boolean` — Optional. If true, actions will not be served for the request. This only works for enterprise edition.
- **query** `Query` — Optional. Current user query. Empty query is only supported if `file_ids` are provided. In this case, the answer will be generated based on those context files.
  - **queryId** `string` — Output only. Unique Id for the query.
  - **parts** `array<QueryPart>` — Query content parts.
    - **uiJsonPayload** `string` — This field is expected to be a ui message in JSON format. As of Q1 2026, ui_json_payload is only supported for A2UI messages.
    - **mimeType** `string` — Optional. The IANA standard MIME type of the data. See https://www.iana.org/assignments/media-types/media-types.xhtml. This field is optional. If not set, the default assumed MI…
    - **text** `string` — Text content.
    - **driveDocumentReference** `QueryPartDriveDocumentReference` — Reference to a Google Drive document.
      - **displayTitle** `string` — The display title of the reference.
      - **documentName** `string` — The full resource name of the document. Format: `projects/*/locations/*/collections/*/dataStores/*/branches/*/documents/*`.
      - **fileId** `string` — Output only. The file id of the Drive document data stored in the session context files.
      - **iconUri** `string` — The icon uri of the Drive document reference.
      - **destinationUri** `string` — The destination uri of the reference.
      - **driveId** `string` — The Drive id of the document.
    - **personReference** `QueryPartPersonReference` — Reference to a person.
      - **displayPhotoUri** `string` — The display photo url of the person.
      - **displayName** `string` — The display name of the person.
      - **documentName** `string` — The full resource name of the person. Format: `projects/*/locations/*/collections/*/dataStores/*/branches/*/documents/*`.
      - **fileId** `string` — Output only. The file id of the person data stored in the session context files.
      - **destinationUri** `string` — The destination uri of the person.
      - **email** `string` — The email of the person.
      - **personId** `string` — The person id of the person.
    - **documentReference** `QueryPartDocumentReference` — Other VAIS Document references.
      - **urlForConnector** `string` — Input only. The url_for_connector of the document returned by Federated Search.
      - **displayTitle** `string` — The display title of the reference.
      - **documentName** `string` — The full resource name of the document. Format: `projects/{project}/locations/{location}/collections/{collection}/dataStores/{data_store}/branches/{branch}/documents/{document_i…
      - **fileId** `string` — Output only. The file id of the document data stored in the session context files.
      - **destinationUri** `string` — The destination uri of the reference.
      - **iconUri** `string` — The icon uri of the reference.
  - **text** `string` — Plain text.
  - **createTime** `string` — Output only. The time at which the server accepted this query.
- **userMetadata** `AssistUserMetadata` — Optional. Information about the user initiating the query.
  - **preferredLanguageCode** `string` — Optional. Preferred language to be used for answering if language detection fails. Also used as the language of error messages created by actions, regardless of language detecti…
  - **timeZone** `string` — Optional. IANA time zone, e.g. Europe/Budapest.
- **toolsSpec** `StreamAssistRequestToolsSpec` — Optional. Specification of tools that are used to serve the request.
  - **vertexAiSearchSpec** `StreamAssistRequestToolsSpecVertexAiSearchSpec` — Optional. Specification of the Vertex AI Search tool.
    - **dataStoreSpecs** `array<SearchRequestDataStoreSpec>` — Optional. Specs defining DataStores to filter on in a search call and configurations for those data stores. This is only considered for Engines with multiple data stores.
      - **numResults** `integer` — Optional. The maximum number of results to retrieve from this data store. If not specified, it will use the SearchRequest.num_results_per_data_store if provided, otherwise there…
      - **boostSpec** `SearchRequestBoostSpec` — Optional. Boost specification to boost certain documents. For more information on boosting, see [Boosting](https://cloud.google.com/generative-ai-app-builder/docs/boost-search-r…
        - **conditionBoostSpecs** `array<SearchRequestBoostSpecConditionBoostSpec>` — Condition boost specifications. If a document matches multiple conditions in the specifications, boost scores from these specifications are all applied and combined in a non-lin…
          - **condition** `string` — An expression which specifies a boost condition. The syntax and supported fields are the same as a filter expression. See SearchRequest.filter for detail syntax and limitations.…
          - **boost** `number` — Strength of the condition boost, which should be in [-1, 1]. Negative boost means demotion. Default is 0.0. Setting to 1.0 gives the document a big promotion. However, it does n…
          - **boostControlSpec** `SearchRequestBoostSpecConditionBoostSpecBoostControlSpec` — Complex specification for custom ranking based on customer defined attribute value.
      - **filter** `string` — Optional. Filter specification to filter documents in the data store specified by data_store field. For more information on filtering, see [Filtering](https://cloud.google.com/g…
      - **dataStore** `string` — Required. Full resource name of DataStore, such as `projects/{project}/locations/{location}/collections/{collection_id}/dataStores/{data_store_id}`. The path must include the pr…
      - **customSearchOperators** `string` — Optional. Custom search operators which if specified will be used to filter results from workspace data stores. For more information on custom search operators, see [SearchOpera…
    - **filter** `string` — Optional. The filter syntax consists of an expression language for constructing a predicate from one or more fields of the documents being filtered. Filter expression is case-se…
  - **webGroundingSpec** `StreamAssistRequestToolsSpecWebGroundingSpec` — Optional. Specification of the web grounding tool. If field is present, enables grounding with web search. Works only if Assistant.web_grounding_type is WEB_GROUNDING_TYPE_GOOGL…
  - **videoGenerationSpec** `StreamAssistRequestToolsSpecVideoGenerationSpec` — Optional. Specification of the video generation tool.
  - **imageGenerationSpec** `StreamAssistRequestToolsSpecImageGenerationSpec` — Optional. Specification of the image generation tool.
- **generationSpec** `StreamAssistRequestGenerationSpec` — Optional. Specification of the generation configuration for the request.
  - **modelId** `string` — Optional. The Vertex AI model_id used for the generative model. If not set, the default Assistant model will be used.

### Response — `StreamAssistResponse`

- **sessionInfo** `StreamAssistResponseSessionInfo` — Session information. Only included in the final StreamAssistResponse of the response stream.
  - **session** `string` — Name of the newly generated or continued session. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/sessions/{session}`.
- **invokedSkills** `array<StreamAssistResponseInvokedSkill>` — The skills executed during the turn.
  - **displayName** `string` — The display name of the skill.
  - **name** `string` — The resource name of the skill.
- **invocationTools** `array<string>` — The tool names of the tools that were invoked.
- **answer** `AssistAnswer` — Assist answer resource object containing parts of the assistant's final answer for the user's query. Not present if the current response doesn't add anything to previously sent …
  - **assistSkippedReasons** `array<enum>` — Reasons for not answering the assist call.
  - **replies** `array<AssistAnswerReply>` — Replies of the assistant.
    - **createTime** `string` — The time when the reply was created.
    - **groundedContent** `AssistantGroundedContent` — Possibly grounded response text or media from the assistant.
      - **textGroundingMetadata** `AssistantGroundedContentTextGroundingMetadata` — Metadata for grounding based on text sources.
        - **segments** `array<AssistantGroundedContentTextGroundingMetadataSegment>` — Grounding information for parts of the text.
          - **text** `string` — The text segment itself.
          - **startIndex** `string` — Zero-based index indicating the start of the segment, measured in bytes of a UTF-8 string (i.e. characters encoded on multiple bytes have a length of more than one).
          - **groundingScore** `number` — Score for the segment.
          - **endIndex** `string` — End of the segment, exclusive.
          - **referenceIndices** `array<integer>` — References for the segment.
        - **visualSegments** `array<AssistantGroundedContentTextGroundingMetadataVisualSegment>` — Grounding information for parts of the visual content.
          - **contentId** `string` — The content id of the visual segment. In order to display the citation of the visual element, this content_id needs to match with the `grounded_content.content_metadata.content_…
          - **referenceIndices** `array<integer>` — References for the visual segment.
        - **references** `array<AssistantGroundedContentTextGroundingMetadataReference>` — References for the grounded text.
          - **codeSnippet** `string` — Chunk of code snippet from the referenced document.
          - **content** `string` — Referenced text content.
          - **documentMetadata** `AssistantGroundedContentTextGroundingMetadataReferenceDocumentMetadata` — Document metadata.
      - **content** `AssistantContent` — The content.
        - **thought** `boolean` — Optional. Indicates if the part is thought from the model.
        - **file** `AssistantContentFile` — A file, e.g., an audio summary.
          - **mimeType** `string` — Required. The media type (MIME type) of the file.
          - **fileId** `string` — Required. The file ID.
        - **text** `string` — Inline text.
        - **codeExecutionResult** `AssistantContentCodeExecutionResult` — Result of executing an ExecutableCode.
          - **output** `string` — Optional. Contains stdout when code execution is successful, stderr or other description otherwise.
          - **outcome** `enum` — enum: OUTCOME_UNSPECIFIED, OUTCOME_OK, OUTCOME_FAILED, OUTCOME_DEADLINE_EXCEEDED — Required. Outcome of the code execution.
        - **role** `string` — The producer of the content. Can be "model" or "user".
        - **inlineData** `AssistantContentBlob` — Inline binary data.
          - **mimeType** `string` — Required. The media type (MIME type) of the generated data.
          - **data** `string` — Required. Raw bytes.
        - **executableCode** `AssistantContentExecutableCode` — Code generated by the model that is meant to be executed.
          - **code** `string` — Required. The code content. Currently only supports Python.
      - **citationMetadata** `CitationMetadata` — Source attribution of the generated content. See also https://cloud.google.com/vertex-ai/generative-ai/docs/learn/overview#citation_check
        - **citations** `array<Citation>` — Output only. List of citations.
          - **uri** `string` — Output only. Url reference of the attribution.
          - **license** `string` — Output only. License of the attribution.
          - **startIndex** `integer` — Output only. Start index into the content.
          - **endIndex** `integer` — Output only. End index into the content.
          - **title** `string` — Output only. Title of the attribution.
          - **publicationDate** `GoogleTypeDate` — Output only. Publication date of the attribution.
    - **replyId** `string` — Output only. When set, uniquely identifies a reply within the `AssistAnswer` resource. During an AssistantService.StreamAssist call, multiple `Reply` messages with the same ID c…
  - **name** `string` — Immutable. Identifier. Resource name of the `AssistAnswer`. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/sessions/{session}/assistA…
  - **customerPolicyEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResult` — Optional. The field contains information about the various policy checks' results like the banned phrases or the Model Armor checks. This field is populated only if the assist c…
    - **verdict** `enum` — enum: UNSPECIFIED, ALLOW, BLOCK — Final verdict of the customer policy enforcement. If only one policy blocked the processing, the verdict is BLOCK.
    - **policyResults** `array<AssistAnswerCustomerPolicyEnforcementResultPolicyEnforcementResult>` — Customer policy enforcement results. Populated only if the assist call was skipped due to a policy violation. It contains results from those filters that blocked the processing …
      - **modelArmorEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResultModelArmorEnforcementResult` — The policy enforcement result for the Model Armor policy.
        - **error** `GoogleRpcStatus` — The error returned by Model Armor if the policy enforcement failed for some reason.
          - **message** `string` — A developer-facing error message, which should be in English. Any user-facing error message should be localized and sent in the google.rpc.Status.details field, or localized by …
          - **code** `integer` — The status code, which should be an enum value of google.rpc.Code.
          - **details** `array<object>` — A list of messages that carry the error details. There is a common set of message types for APIs to use.
        - **modelArmorViolation** `string` — The Model Armor violation that was found.
      - **bannedPhraseEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResultBannedPhraseEnforcementResult` — The policy enforcement result for the banned phrase policy.
        - **bannedPhrases** `array<string>` — The banned phrases that were found in the query or the answer.
  - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, SKIPPED, CANCELLED — State of the answer generation.
- **assistToken** `string` — A global unique ID that identifies the current pair of request and stream of responses. Used for feedback and support.

