# Assistant & engine configuration (Google-side, reference only)

These are configured by the Gemini Enterprise admin, **not** by the add-in, but the client may read them (e.g. to show enabled tools or canned queries). Includes Model Armor and banned-phrase policy — confirming guardrails live in the engine config, per the architecture.

## assistants.get


Read an assistant config.

- **HTTP**: `GET` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/collections/{collectionsId}/engines/{enginesId}/assistants/{assistantsId}`
- **Method id**: `projects.locations.collections.engines.assistants.get`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.assist.readwrite, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

### Response — `Assistant`

- **disableLocationContext** `boolean` — Optional. Indicates whether to disable user location context. By default, user location context is enabled.
- **updateTime** `string` — Output only. Represents the time when this Assistant was most recently updated.
- **enabledTools** `object` — Optional. Note: not implemented yet. Use enabled_actions instead. The enabled tools on this assistant. The keys are connector name, for example "projects/{projectId}/locations/{…
- **displayName** `string` — Required. The assistant display name. It must be a UTF-8 encoded string with a length limit of 128 characters.
- **webGroundingType** `enum` — enum: WEB_GROUNDING_TYPE_UNSPECIFIED, WEB_GROUNDING_TYPE_DISABLED, WEB_GROUNDING_TYPE_GOOGLE_SEARCH, WEB_GROUNDING_TYPE_ENTERPRISE_WEB_SEARCH — Optional. The type of web grounding to use.
- **description** `string` — Optional. Description for additional information. Expected to be shown on the configuration UI, not to the users of the assistant.
- **name** `string` — Immutable. Resource name of the assistant. Format: `projects/{project}/locations/{location}/collections/{collection}/engines/{engine}/assistants/{assistant}` It must be a UTF-8 …
- **defaultWebGroundingToggleOff** `boolean` — Optional. This field controls the default web grounding toggle for end users if `web_grounding_type` is set to `WEB_GROUNDING_TYPE_GOOGLE_SEARCH` or `WEB_GROUNDING_TYPE_ENTERPRI…
- **customerPolicy** `AssistantCustomerPolicy` — Optional. Customer policy for the assistant.
  - **bannedPhrases** `array<AssistantCustomerPolicyBannedPhrase>` — Optional. List of banned phrases.
    - **phrase** `string` — Required. The raw string content to be banned.
    - **matchType** `enum` — enum: BANNED_PHRASE_MATCH_TYPE_UNSPECIFIED, SIMPLE_STRING_MATCH, WORD_BOUNDARY_STRING_MATCH — Optional. Match type for the banned phrase.
    - **ignoreDiacritics** `boolean` — Optional. If true, diacritical marks (e.g., accents, umlauts) are ignored when matching banned phrases. For example, "cafe" would match "café".
  - **modelArmorConfig** `AssistantCustomerPolicyModelArmorConfig` — Optional. Model Armor configuration to be used for sanitizing user prompts and assistant responses.
    - **userPromptTemplate** `string` — Optional. The resource name of the Model Armor template for sanitizing user prompts. Format: `projects/{project}/locations/{location}/templates/{template_id}` If not specified, …
    - **failureMode** `enum` — enum: FAILURE_MODE_UNSPECIFIED, FAIL_OPEN, FAIL_CLOSED — Optional. Defines the failure mode for Model Armor sanitization.
    - **responseTemplate** `string` — Optional. The resource name of the Model Armor template for sanitizing assistant responses. Format: `projects/{project}/locations/{location}/templates/{template_id}` If not spec…
  - **dataProtectionPolicy** `DataProtectionPolicy` — Optional. Data protection policy to be used for sanitizing file uploads.
    - **sensitiveDataProtectionPolicy** `DataProtectionPolicySensitiveDataProtectionPolicy` — Optional. Specifies the sensitive data protection policy for the connector source.
      - **policy** `string` — Optional. Specifies the resource name of the Sensitive Data Protection content policy.
- **createTime** `string` — Output only. Represents the time when this Assistant was created.
- **generationConfig** `AssistantGenerationConfig` — Optional. Configuration for the generation of the assistant response.
  - **defaultModelId** `string` — Optional. The default model to use for assistant.
  - **allowedModelIds** `array<string>` — Optional. The list of models that are allowed to be used for assistant.
  - **systemInstruction** `AssistantGenerationConfigSystemInstruction` — System instruction, also known as the prompt preamble for LLM calls. See also https://cloud.google.com/vertex-ai/generative-ai/docs/learn/prompts/system-instructions
    - **additionalSystemInstruction** `string` — Optional. Additional system instruction that will be added to the default system instruction.
  - **defaultLanguage** `string` — The default language to use for the generation of the assistant response. Use an ISO 639-1 language code such as `en`. If not specified, the language will be automatically detec…

