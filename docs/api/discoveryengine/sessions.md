# Sessions — multi-turn conversation state

Sessions persist conversation context server-side so a reopened document can resume. Create a session, pass its name to streamAssist/answer, and persist the session id in the host metadata (provenance `sessionId`).

## sessions.create


Create a conversation session.

- **HTTP**: `POST` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/collections/{collectionsId}/engines/{enginesId}/sessions`
- **Method id**: `projects.locations.collections.engines.sessions.create`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.assist.readwrite, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

- **Query params**: `sessionId`

### Request body — `Session`

- **startTime** `string` — Output only. The time the session started.
- **labels** `array<string>` — Optional. The labels for the session. Can be set as filter in ListSessionsRequest.
- **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*`
- **isPinned** `boolean` — Optional. Whether the session is pinned, pinned session will be displayed on the top of the session list.
- **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS — The state of the session.
- **userPseudoId** `string` — A unique identifier for tracking users.
- **pendingAsyncAssistOperationId** `string` — Output only. Full resource name of an in-progress AsyncAssist operation for this session, e.g. `projects/*/locations/*/collections/*/engines/*/sessions/*/operations/*`. Set when…
- **endTime** `string` — Output only. The time the session finished.
- **turns** `array<SessionTurn>` — Turns.
  - **live** `boolean` — Optional. Indicates whether this turn is a live turn.
  - **query** `Query` — Optional. The user query. May not be set if this turn is merely regenerating an answer to a different turn
    - **queryId** `string` — Output only. Unique Id for the query.
    - **parts** `array<QueryPart>` — Query content parts.
      - **uiJsonPayload** `string` — This field is expected to be a ui message in JSON format. As of Q1 2026, ui_json_payload is only supported for A2UI messages.
      - **mimeType** `string` — Optional. The IANA standard MIME type of the data. See https://www.iana.org/assignments/media-types/media-types.xhtml. This field is optional. If not set, the default assumed MI…
      - **text** `string` — Text content.
      - **driveDocumentReference** `QueryPartDriveDocumentReference` — Reference to a Google Drive document.
      - **personReference** `QueryPartPersonReference` — Reference to a person.
      - **documentReference** `QueryPartDocumentReference` — Other VAIS Document references.
    - **text** `string` — Plain text.
    - **createTime** `string` — Output only. The time at which the server accepted this query.
  - **answer** `string` — Optional. The resource name of the answer to the user query. Only set if the answer generation (/answer API call) happened in this turn.
  - **detailedAssistAnswer** `AssistAnswer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting assistant session.
    - **assistSkippedReasons** `array<enum>` — Reasons for not answering the assist call.
    - **replies** `array<AssistAnswerReply>` — Replies of the assistant.
      - **createTime** `string` — The time when the reply was created.
      - **groundedContent** `AssistantGroundedContent` — Possibly grounded response text or media from the assistant.
      - **replyId** `string` — Output only. When set, uniquely identifies a reply within the `AssistAnswer` resource. During an AssistantService.StreamAssist call, multiple `Reply` messages with the same ID c…
    - **name** `string` — Immutable. Identifier. Resource name of the `AssistAnswer`. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/sessions/{session}/assistA…
    - **customerPolicyEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResult` — Optional. The field contains information about the various policy checks' results like the banned phrases or the Model Armor checks. This field is populated only if the assist c…
      - **verdict** `enum` — enum: UNSPECIFIED, ALLOW, BLOCK — Final verdict of the customer policy enforcement. If only one policy blocked the processing, the verdict is BLOCK.
      - **policyResults** `array<AssistAnswerCustomerPolicyEnforcementResultPolicyEnforcementResult>` — Customer policy enforcement results. Populated only if the assist call was skipped due to a policy violation. It contains results from those filters that blocked the processing …
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, SKIPPED, CANCELLED — State of the answer generation.
  - **detailedAnswer** `Answer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting answer query se…
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, STREAMING — The state of the answer generation.
    - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded the answer is by the reference chunks.
    - **createTime** `string` — Output only. Answer creation timestamp.
    - **safetyRatings** `array<SafetyRating>` — Optional. Safety ratings.
      - **category** `enum` — enum: HARM_CATEGORY_UNSPECIFIED, HARM_CATEGORY_HATE_SPEECH, HARM_CATEGORY_DANGEROUS_CONTENT, HARM_CATEGORY_HARASSMENT, HARM_CATEGORY_SEXUALLY_EXPLICIT, HARM_CATEGORY_CIVIC_INTEGRITY — Output only. Harm category.
      - **severity** `enum` — enum: HARM_SEVERITY_UNSPECIFIED, HARM_SEVERITY_NEGLIGIBLE, HARM_SEVERITY_LOW, HARM_SEVERITY_MEDIUM, HARM_SEVERITY_HIGH — Output only. Harm severity levels in the content.
      - **severityScore** `number` — Output only. Harm severity score.
      - **blocked** `boolean` — Output only. Indicates whether the content was filtered out because of this rating.
      - **probability** `enum` — enum: HARM_PROBABILITY_UNSPECIFIED, NEGLIGIBLE, LOW, MEDIUM, HIGH — Output only. Harm probability levels in the content.
      - **probabilityScore** `number` — Output only. Harm probability score.
    - **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*/answers/*`
    - **references** `array<AnswerReference>` — References.
      - **structuredDocumentInfo** `AnswerReferenceStructuredDocumentInfo` — Structured document information.
      - **unstructuredDocumentInfo** `AnswerReferenceUnstructuredDocumentInfo` — Unstructured document information.
      - **chunkInfo** `AnswerReferenceChunkInfo` — Chunk information.
    - **groundingSupports** `array<AnswerGroundingSupport>` — Optional. Grounding supports.
      - **startIndex** `string` — Required. Index indicates the start of the claim, measured in bytes (UTF-8 unicode).
      - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded is a specific claim by the references. Higher value means that the claim is better supported by the reference chunks.
      - **groundingCheckRequired** `boolean` — Indicates that this claim required grounding check. When the system decided this claim didn't require attribution/grounding check, this field is set to false. In that case, no g…
      - **sources** `array<AnswerCitationSource>` — Optional. Citation sources for the claim.
      - **endIndex** `string` — Required. End of the claim, exclusive.
    - **answerSkippedReasons** `array<enum>` — Additional answer-skipped reasons. This provides the reason for ignored cases. If nothing is skipped, this field is not set.
    - **relatedQuestions** `array<string>` — Suggested related questions.
    - **completeTime** `string` — Output only. Answer completed timestamp.
    - **answerText** `string` — The textual answer.
    - **blobAttachments** `array<AnswerBlobAttachment>` — Output only. List of blob attachments in the answer.
      - **data** `AnswerBlobAttachmentBlob` — Output only. The mime type and data of the blob.
      - **attributionType** `enum` — enum: ATTRIBUTION_TYPE_UNSPECIFIED, CORPUS, GENERATED — Output only. The attribution type of the blob.
    - **steps** `array<AnswerStep>` — Answer generation steps.
      - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED — The state of the step.
      - **thought** `string` — The thought of the step.
      - **actions** `array<AnswerStepAction>` — Actions.
      - **description** `string` — The description of the step.
    - **queryUnderstandingInfo** `AnswerQueryUnderstandingInfo` — Query understanding information.
      - **queryClassificationInfo** `array<AnswerQueryUnderstandingInfoQueryClassificationInfo>` — Query classification information.
    - **citations** `array<AnswerCitation>` — Citations.
      - **startIndex** `string` — Index indicates the start of the segment, measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
      - **sources** `array<AnswerCitationSource>` — Citation sources for the attributed segment.
      - **endIndex** `string` — End of the attributed segment, exclusive. Measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
  - **queryConfig** `object` — Optional. Represents metadata related to the query config, for example LLM model and version used, model parameters (temperature, grounding parameters, etc.). The prefix "google…
- **displayName** `string` — Optional. The display name of the session. This field is used to identify the session in the UI. By default, the display name is the first turn query text in the session.

### Response — `Session`

- **startTime** `string` — Output only. The time the session started.
- **labels** `array<string>` — Optional. The labels for the session. Can be set as filter in ListSessionsRequest.
- **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*`
- **isPinned** `boolean` — Optional. Whether the session is pinned, pinned session will be displayed on the top of the session list.
- **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS — The state of the session.
- **userPseudoId** `string` — A unique identifier for tracking users.
- **pendingAsyncAssistOperationId** `string` — Output only. Full resource name of an in-progress AsyncAssist operation for this session, e.g. `projects/*/locations/*/collections/*/engines/*/sessions/*/operations/*`. Set when…
- **endTime** `string` — Output only. The time the session finished.
- **turns** `array<SessionTurn>` — Turns.
  - **live** `boolean` — Optional. Indicates whether this turn is a live turn.
  - **query** `Query` — Optional. The user query. May not be set if this turn is merely regenerating an answer to a different turn
    - **queryId** `string` — Output only. Unique Id for the query.
    - **parts** `array<QueryPart>` — Query content parts.
      - **uiJsonPayload** `string` — This field is expected to be a ui message in JSON format. As of Q1 2026, ui_json_payload is only supported for A2UI messages.
      - **mimeType** `string` — Optional. The IANA standard MIME type of the data. See https://www.iana.org/assignments/media-types/media-types.xhtml. This field is optional. If not set, the default assumed MI…
      - **text** `string` — Text content.
      - **driveDocumentReference** `QueryPartDriveDocumentReference` — Reference to a Google Drive document.
      - **personReference** `QueryPartPersonReference` — Reference to a person.
      - **documentReference** `QueryPartDocumentReference` — Other VAIS Document references.
    - **text** `string` — Plain text.
    - **createTime** `string` — Output only. The time at which the server accepted this query.
  - **answer** `string` — Optional. The resource name of the answer to the user query. Only set if the answer generation (/answer API call) happened in this turn.
  - **detailedAssistAnswer** `AssistAnswer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting assistant session.
    - **assistSkippedReasons** `array<enum>` — Reasons for not answering the assist call.
    - **replies** `array<AssistAnswerReply>` — Replies of the assistant.
      - **createTime** `string` — The time when the reply was created.
      - **groundedContent** `AssistantGroundedContent` — Possibly grounded response text or media from the assistant.
      - **replyId** `string` — Output only. When set, uniquely identifies a reply within the `AssistAnswer` resource. During an AssistantService.StreamAssist call, multiple `Reply` messages with the same ID c…
    - **name** `string` — Immutable. Identifier. Resource name of the `AssistAnswer`. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/sessions/{session}/assistA…
    - **customerPolicyEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResult` — Optional. The field contains information about the various policy checks' results like the banned phrases or the Model Armor checks. This field is populated only if the assist c…
      - **verdict** `enum` — enum: UNSPECIFIED, ALLOW, BLOCK — Final verdict of the customer policy enforcement. If only one policy blocked the processing, the verdict is BLOCK.
      - **policyResults** `array<AssistAnswerCustomerPolicyEnforcementResultPolicyEnforcementResult>` — Customer policy enforcement results. Populated only if the assist call was skipped due to a policy violation. It contains results from those filters that blocked the processing …
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, SKIPPED, CANCELLED — State of the answer generation.
  - **detailedAnswer** `Answer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting answer query se…
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, STREAMING — The state of the answer generation.
    - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded the answer is by the reference chunks.
    - **createTime** `string` — Output only. Answer creation timestamp.
    - **safetyRatings** `array<SafetyRating>` — Optional. Safety ratings.
      - **category** `enum` — enum: HARM_CATEGORY_UNSPECIFIED, HARM_CATEGORY_HATE_SPEECH, HARM_CATEGORY_DANGEROUS_CONTENT, HARM_CATEGORY_HARASSMENT, HARM_CATEGORY_SEXUALLY_EXPLICIT, HARM_CATEGORY_CIVIC_INTEGRITY — Output only. Harm category.
      - **severity** `enum` — enum: HARM_SEVERITY_UNSPECIFIED, HARM_SEVERITY_NEGLIGIBLE, HARM_SEVERITY_LOW, HARM_SEVERITY_MEDIUM, HARM_SEVERITY_HIGH — Output only. Harm severity levels in the content.
      - **severityScore** `number` — Output only. Harm severity score.
      - **blocked** `boolean` — Output only. Indicates whether the content was filtered out because of this rating.
      - **probability** `enum` — enum: HARM_PROBABILITY_UNSPECIFIED, NEGLIGIBLE, LOW, MEDIUM, HIGH — Output only. Harm probability levels in the content.
      - **probabilityScore** `number` — Output only. Harm probability score.
    - **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*/answers/*`
    - **references** `array<AnswerReference>` — References.
      - **structuredDocumentInfo** `AnswerReferenceStructuredDocumentInfo` — Structured document information.
      - **unstructuredDocumentInfo** `AnswerReferenceUnstructuredDocumentInfo` — Unstructured document information.
      - **chunkInfo** `AnswerReferenceChunkInfo` — Chunk information.
    - **groundingSupports** `array<AnswerGroundingSupport>` — Optional. Grounding supports.
      - **startIndex** `string` — Required. Index indicates the start of the claim, measured in bytes (UTF-8 unicode).
      - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded is a specific claim by the references. Higher value means that the claim is better supported by the reference chunks.
      - **groundingCheckRequired** `boolean` — Indicates that this claim required grounding check. When the system decided this claim didn't require attribution/grounding check, this field is set to false. In that case, no g…
      - **sources** `array<AnswerCitationSource>` — Optional. Citation sources for the claim.
      - **endIndex** `string` — Required. End of the claim, exclusive.
    - **answerSkippedReasons** `array<enum>` — Additional answer-skipped reasons. This provides the reason for ignored cases. If nothing is skipped, this field is not set.
    - **relatedQuestions** `array<string>` — Suggested related questions.
    - **completeTime** `string` — Output only. Answer completed timestamp.
    - **answerText** `string` — The textual answer.
    - **blobAttachments** `array<AnswerBlobAttachment>` — Output only. List of blob attachments in the answer.
      - **data** `AnswerBlobAttachmentBlob` — Output only. The mime type and data of the blob.
      - **attributionType** `enum` — enum: ATTRIBUTION_TYPE_UNSPECIFIED, CORPUS, GENERATED — Output only. The attribution type of the blob.
    - **steps** `array<AnswerStep>` — Answer generation steps.
      - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED — The state of the step.
      - **thought** `string` — The thought of the step.
      - **actions** `array<AnswerStepAction>` — Actions.
      - **description** `string` — The description of the step.
    - **queryUnderstandingInfo** `AnswerQueryUnderstandingInfo` — Query understanding information.
      - **queryClassificationInfo** `array<AnswerQueryUnderstandingInfoQueryClassificationInfo>` — Query classification information.
    - **citations** `array<AnswerCitation>` — Citations.
      - **startIndex** `string` — Index indicates the start of the segment, measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
      - **sources** `array<AnswerCitationSource>` — Citation sources for the attributed segment.
      - **endIndex** `string` — End of the attributed segment, exclusive. Measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
  - **queryConfig** `object` — Optional. Represents metadata related to the query config, for example LLM model and version used, model parameters (temperature, grounding parameters, etc.). The prefix "google…
- **displayName** `string` — Optional. The display name of the session. This field is used to identify the session in the UI. By default, the display name is the first turn query text in the session.


## sessions.get


Fetch a session (history of turns).

- **HTTP**: `GET` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/collections/{collectionsId}/engines/{enginesId}/sessions/{sessionsId}`
- **Method id**: `projects.locations.collections.engines.sessions.get`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.assist.readwrite, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

- **Query params**: `includeAnswerDetails`

### Response — `Session`

- **startTime** `string` — Output only. The time the session started.
- **labels** `array<string>` — Optional. The labels for the session. Can be set as filter in ListSessionsRequest.
- **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*`
- **isPinned** `boolean` — Optional. Whether the session is pinned, pinned session will be displayed on the top of the session list.
- **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS — The state of the session.
- **userPseudoId** `string` — A unique identifier for tracking users.
- **pendingAsyncAssistOperationId** `string` — Output only. Full resource name of an in-progress AsyncAssist operation for this session, e.g. `projects/*/locations/*/collections/*/engines/*/sessions/*/operations/*`. Set when…
- **endTime** `string` — Output only. The time the session finished.
- **turns** `array<SessionTurn>` — Turns.
  - **live** `boolean` — Optional. Indicates whether this turn is a live turn.
  - **query** `Query` — Optional. The user query. May not be set if this turn is merely regenerating an answer to a different turn
    - **queryId** `string` — Output only. Unique Id for the query.
    - **parts** `array<QueryPart>` — Query content parts.
      - **uiJsonPayload** `string` — This field is expected to be a ui message in JSON format. As of Q1 2026, ui_json_payload is only supported for A2UI messages.
      - **mimeType** `string` — Optional. The IANA standard MIME type of the data. See https://www.iana.org/assignments/media-types/media-types.xhtml. This field is optional. If not set, the default assumed MI…
      - **text** `string` — Text content.
      - **driveDocumentReference** `QueryPartDriveDocumentReference` — Reference to a Google Drive document.
      - **personReference** `QueryPartPersonReference` — Reference to a person.
      - **documentReference** `QueryPartDocumentReference` — Other VAIS Document references.
    - **text** `string` — Plain text.
    - **createTime** `string` — Output only. The time at which the server accepted this query.
  - **answer** `string` — Optional. The resource name of the answer to the user query. Only set if the answer generation (/answer API call) happened in this turn.
  - **detailedAssistAnswer** `AssistAnswer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting assistant session.
    - **assistSkippedReasons** `array<enum>` — Reasons for not answering the assist call.
    - **replies** `array<AssistAnswerReply>` — Replies of the assistant.
      - **createTime** `string` — The time when the reply was created.
      - **groundedContent** `AssistantGroundedContent` — Possibly grounded response text or media from the assistant.
      - **replyId** `string` — Output only. When set, uniquely identifies a reply within the `AssistAnswer` resource. During an AssistantService.StreamAssist call, multiple `Reply` messages with the same ID c…
    - **name** `string` — Immutable. Identifier. Resource name of the `AssistAnswer`. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/sessions/{session}/assistA…
    - **customerPolicyEnforcementResult** `AssistAnswerCustomerPolicyEnforcementResult` — Optional. The field contains information about the various policy checks' results like the banned phrases or the Model Armor checks. This field is populated only if the assist c…
      - **verdict** `enum` — enum: UNSPECIFIED, ALLOW, BLOCK — Final verdict of the customer policy enforcement. If only one policy blocked the processing, the verdict is BLOCK.
      - **policyResults** `array<AssistAnswerCustomerPolicyEnforcementResultPolicyEnforcementResult>` — Customer policy enforcement results. Populated only if the assist call was skipped due to a policy violation. It contains results from those filters that blocked the processing …
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, SKIPPED, CANCELLED — State of the answer generation.
  - **detailedAnswer** `Answer` — Output only. In ConversationalSearchService.GetSession API, if GetSessionRequest.include_answer_details is set to true, this field will be populated when getting answer query se…
    - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED, STREAMING — The state of the answer generation.
    - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded the answer is by the reference chunks.
    - **createTime** `string` — Output only. Answer creation timestamp.
    - **safetyRatings** `array<SafetyRating>` — Optional. Safety ratings.
      - **category** `enum` — enum: HARM_CATEGORY_UNSPECIFIED, HARM_CATEGORY_HATE_SPEECH, HARM_CATEGORY_DANGEROUS_CONTENT, HARM_CATEGORY_HARASSMENT, HARM_CATEGORY_SEXUALLY_EXPLICIT, HARM_CATEGORY_CIVIC_INTEGRITY — Output only. Harm category.
      - **severity** `enum` — enum: HARM_SEVERITY_UNSPECIFIED, HARM_SEVERITY_NEGLIGIBLE, HARM_SEVERITY_LOW, HARM_SEVERITY_MEDIUM, HARM_SEVERITY_HIGH — Output only. Harm severity levels in the content.
      - **severityScore** `number` — Output only. Harm severity score.
      - **blocked** `boolean` — Output only. Indicates whether the content was filtered out because of this rating.
      - **probability** `enum` — enum: HARM_PROBABILITY_UNSPECIFIED, NEGLIGIBLE, LOW, MEDIUM, HIGH — Output only. Harm probability levels in the content.
      - **probabilityScore** `number` — Output only. Harm probability score.
    - **name** `string` — Immutable. Fully qualified name `projects/{project}/locations/global/collections/{collection}/engines/{engine}/sessions/*/answers/*`
    - **references** `array<AnswerReference>` — References.
      - **structuredDocumentInfo** `AnswerReferenceStructuredDocumentInfo` — Structured document information.
      - **unstructuredDocumentInfo** `AnswerReferenceUnstructuredDocumentInfo` — Unstructured document information.
      - **chunkInfo** `AnswerReferenceChunkInfo` — Chunk information.
    - **groundingSupports** `array<AnswerGroundingSupport>` — Optional. Grounding supports.
      - **startIndex** `string` — Required. Index indicates the start of the claim, measured in bytes (UTF-8 unicode).
      - **groundingScore** `number` — A score in the range of [0, 1] describing how grounded is a specific claim by the references. Higher value means that the claim is better supported by the reference chunks.
      - **groundingCheckRequired** `boolean` — Indicates that this claim required grounding check. When the system decided this claim didn't require attribution/grounding check, this field is set to false. In that case, no g…
      - **sources** `array<AnswerCitationSource>` — Optional. Citation sources for the claim.
      - **endIndex** `string` — Required. End of the claim, exclusive.
    - **answerSkippedReasons** `array<enum>` — Additional answer-skipped reasons. This provides the reason for ignored cases. If nothing is skipped, this field is not set.
    - **relatedQuestions** `array<string>` — Suggested related questions.
    - **completeTime** `string` — Output only. Answer completed timestamp.
    - **answerText** `string` — The textual answer.
    - **blobAttachments** `array<AnswerBlobAttachment>` — Output only. List of blob attachments in the answer.
      - **data** `AnswerBlobAttachmentBlob` — Output only. The mime type and data of the blob.
      - **attributionType** `enum` — enum: ATTRIBUTION_TYPE_UNSPECIFIED, CORPUS, GENERATED — Output only. The attribution type of the blob.
    - **steps** `array<AnswerStep>` — Answer generation steps.
      - **state** `enum` — enum: STATE_UNSPECIFIED, IN_PROGRESS, FAILED, SUCCEEDED — The state of the step.
      - **thought** `string` — The thought of the step.
      - **actions** `array<AnswerStepAction>` — Actions.
      - **description** `string` — The description of the step.
    - **queryUnderstandingInfo** `AnswerQueryUnderstandingInfo` — Query understanding information.
      - **queryClassificationInfo** `array<AnswerQueryUnderstandingInfoQueryClassificationInfo>` — Query classification information.
    - **citations** `array<AnswerCitation>` — Citations.
      - **startIndex** `string` — Index indicates the start of the segment, measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
      - **sources** `array<AnswerCitationSource>` — Citation sources for the attributed segment.
      - **endIndex** `string` — End of the attributed segment, exclusive. Measured in bytes (UTF-8 unicode). If there are multi-byte characters,such as non-ASCII characters, the index measurement is longer tha…
  - **queryConfig** `object` — Optional. Represents metadata related to the query config, for example LLM model and version used, model parameters (temperature, grounding parameters, etc.). The prefix "google…
- **displayName** `string` — Optional. The display name of the session. This field is used to identify the session in the UI. By default, the display name is the first turn query text in the session.

