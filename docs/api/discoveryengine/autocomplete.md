# completeQuery — query autocomplete / suggestions

Powers type-ahead in the panel composer and Excel formula prompts.

## completionConfig.completeQuery


Suggestions for a partial query.

- **HTTP**: `POST` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/collections/{collectionsId}/engines/{enginesId}/completionConfig:completeQuery`
- **Method id**: `projects.locations.collections.engines.completionConfig.completeQuery`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/cloud_search.query, https://www.googleapis.com/auth/discoveryengine.assist.readwrite, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

### Request body — `AdvancedCompleteQueryRequest`

- **includeTailSuggestions** `boolean` — Indicates if tail suggestions should be returned if there are no suggestions that match the full query. Even if set to true, if there are suggestions that match the full query, …
- **suggestionTypes** `array<enum>` — Optional. Suggestion types to return. If empty or unspecified, query suggestions are returned. Only one suggestion type is supported at the moment.
- **queryModel** `string` — Specifies the autocomplete query model, which only applies to the QUERY SuggestionType. This overrides any model specified in the Configuration > Autocomplete section of the Clo…
- **boostSpec** `AdvancedCompleteQueryRequestBoostSpec` — Optional. Specification to boost suggestions matching the condition.
  - **conditionBoostSpecs** `array<AdvancedCompleteQueryRequestBoostSpecConditionBoostSpec>` — Condition boost specifications. If a suggestion matches multiple conditions in the specifications, boost values from these specifications are all applied and combined in a non-l…
    - **condition** `string` — An expression which specifies a boost condition. The syntax is the same as [filter expression syntax](https://cloud.google.com/generative-ai-app-builder/docs/filter-search-metad…
    - **boost** `number` — Strength of the boost, which should be in [-1, 1]. Negative boost means demotion. Default is 0.0. Setting to 1.0 gives the suggestions a big promotion. However, it does not nece…
- **suggestionTypeSpecs** `array<AdvancedCompleteQueryRequestSuggestionTypeSpec>` — Optional. Specification of each suggestion type.
  - **maxSuggestions** `integer` — Optional. Maximum number of suggestions to return for each suggestion type.
  - **suggestionType** `enum` — enum: SUGGESTION_TYPE_UNSPECIFIED, QUERY, PEOPLE, CONTENT, RECENT_SEARCH, GOOGLE_WORKSPACE — Optional. Suggestion type.
- **experimentIds** `array<string>` — Optional. Experiment ids for this request.
- **userPseudoId** `string` — Optional. A unique identifier for tracking visitors. For example, this could be implemented with an HTTP cookie, which should be able to uniquely identify a visitor on a single …
- **query** `string` — Required. The typeahead input used to fetch suggestions. Maximum length is 128 characters. The query can not be empty for most of the suggestion types. If it is empty, an `INVAL…
- **userInfo** `UserInfo` — Optional. Information about the end user. This should be the same identifier information as UserEvent.user_info and SearchRequest.user_info.
  - **userId** `string` — Highly recommended for logged-in users. Unique identifier for logged-in user, such as a user name. Don't set for anonymous users. Always use a hashed value for this ID. Don't se…
  - **userAgent** `string` — User agent as included in the HTTP header. The field must be a UTF-8 encoded string with a length limit of 1,000 characters. Otherwise, an `INVALID_ARGUMENT` error is returned. …
  - **timeZone** `string` — Optional. IANA time zone, e.g. Europe/Budapest.
  - **preciseLocation** `UserInfoPreciseLocation` — Optional. Input only. Precise location of the user. It is used in Custom Ranking to calculate the distance between the user and the relevant documents.
    - **point** `GoogleTypeLatLng` — Optional. Location represented by a latitude/longitude point.
      - **latitude** `number` — The latitude in degrees. It must be in the range [-90.0, +90.0].
      - **longitude** `number` — The longitude in degrees. It must be in the range [-180.0, +180.0].
    - **address** `string` — Optional. Location represented by a natural language address. Will later be geocoded and converted to either a point or a polygon.

### Response — `AdvancedCompleteQueryResponse`

- **contentSuggestions** `array<AdvancedCompleteQueryResponseContentSuggestion>` — Results of the matched content suggestions. The result list is ordered and the first result is the top suggestion.
  - **document** `Document` — The document data snippet in the suggestion. Only a subset of fields will be populated.
    - **name** `string` — Immutable. The full resource name of the document. Format: `projects/{project}/locations/{location}/collections/{collection}/dataStores/{data_store}/branches/{branch}/documents/…
    - **indexTime** `string` — Output only. The time when the document was last indexed. If this field is populated, it means the document has been indexed. While documents typically become searchable within …
    - **derivedStructData** `object` — Output only. This field is OUTPUT_ONLY. It contains derived data that are not in the original input document.
    - **jsonData** `string` — The JSON string representation of the document. It should conform to the registered Schema or an `INVALID_ARGUMENT` error is thrown.
    - **aclInfo** `DocumentAclInfo` — Access control information for the document.
      - **readers** `array<DocumentAclInfoAccessRestriction>` — Readers of the document.
    - **indexStatus** `DocumentIndexStatus` — Output only. The index status of the document. * If document is indexed successfully, the index_time field is populated. * Otherwise, if document is not indexed due to errors, t…
      - **indexTime** `string` — The time when the document was indexed. If this field is populated, it means the document has been indexed. While documents typically become searchable within seconds of indexin…
      - **errorSamples** `array<GoogleRpcStatus>` — A sample of errors encountered while indexing the document. If this field is populated, the document is not indexed due to errors.
      - **pendingMessage** `string` — Immutable. The message indicates the document index is in progress. If this field is populated, the document index is pending.
    - **parentDocumentId** `string` — The identifier of the parent document. Currently supports at most two level document hierarchy. Id should conform to [RFC-1034](https://tools.ietf.org/html/rfc1034) standard wit…
    - **id** `string` — Immutable. The identifier of the document. Id should conform to [RFC-1034](https://tools.ietf.org/html/rfc1034) standard with a length limit of 128 characters.
    - **schemaId** `string` — The identifier of the schema located in the same data store.
    - **content** `DocumentContent` — The unstructured data linked to this document. Content can only be set and must be set if this document is under a `CONTENT_REQUIRED` data store.
      - **rawBytes** `string` — The content represented as a stream of bytes. The maximum length is 1,000,000 bytes (1 MB / ~0.95 MiB). Note: As with all `bytes` fields, this field is represented as pure binar…
      - **mimeType** `string` — The MIME type of the content. Supported types: * `application/pdf` (PDF, only native PDFs are supported for now) * `text/html` (HTML) * `text/plain` (TXT) * `application/xml` or…
      - **uri** `string` — The URI of the content. Only Cloud Storage URIs (e.g. `gs://bucket-name/path/to/file`) are supported. The maximum file size is 2.5 MB for text-based formats, 200 MB for other fo…
    - **structData** `object` — The structured JSON data for the document. It should conform to the registered Schema or an `INVALID_ARGUMENT` error is thrown.
  - **score** `number` — The score of each suggestion. The score is in the range of [0, 1].
  - **destinationUri** `string` — The destination uri of the content suggestion.
  - **contentType** `enum` — enum: CONTENT_TYPE_UNSPECIFIED, GOOGLE_WORKSPACE, THIRD_PARTY — The type of the content suggestion.
  - **iconUri** `string` — The icon uri of the content suggestion.
  - **suggestion** `string` — The suggestion for the query.
  - **dataStore** `string` — The name of the dataStore that this suggestion belongs to.
- **tailMatchTriggered** `boolean` — True if the returned suggestions are all tail suggestions. For tail matching to be triggered, include_tail_suggestions in the request must be true and there must be no suggestio…
- **recentSearchSuggestions** `array<AdvancedCompleteQueryResponseRecentSearchSuggestion>` — Results of the matched "recent search" suggestions. The result list is ordered and the first result is the top suggestion.
  - **suggestion** `string` — The suggestion for the query.
  - **recentSearchTime** `string` — The time when this recent rearch happened.
  - **score** `number` — The score of each suggestion. The score is in the range of [0, 1].
- **querySuggestions** `array<AdvancedCompleteQueryResponseQuerySuggestion>` — Results of the matched query suggestions. The result list is ordered and the first result is a top suggestion.
  - **completableFieldPaths** `array<string>` — The unique document field paths that serve as the source of this suggestion if it was generated from completable fields. This field is only populated for the document-completabl…
  - **score** `number` — The score of each suggestion. The score is in the range of [0, 1].
  - **suggestion** `string` — The suggestion for the query.
  - **dataStore** `array<string>` — The name of the dataStore that this suggestion belongs to.
- **peopleSuggestions** `array<AdvancedCompleteQueryResponsePersonSuggestion>` — Results of the matched people suggestions. The result list is ordered and the first result is the top suggestion.
  - **document** `Document` — The document data snippet in the suggestion. Only a subset of fields is populated.
    - **name** `string` — Immutable. The full resource name of the document. Format: `projects/{project}/locations/{location}/collections/{collection}/dataStores/{data_store}/branches/{branch}/documents/…
    - **indexTime** `string` — Output only. The time when the document was last indexed. If this field is populated, it means the document has been indexed. While documents typically become searchable within …
    - **derivedStructData** `object` — Output only. This field is OUTPUT_ONLY. It contains derived data that are not in the original input document.
    - **jsonData** `string` — The JSON string representation of the document. It should conform to the registered Schema or an `INVALID_ARGUMENT` error is thrown.
    - **aclInfo** `DocumentAclInfo` — Access control information for the document.
      - **readers** `array<DocumentAclInfoAccessRestriction>` — Readers of the document.
    - **indexStatus** `DocumentIndexStatus` — Output only. The index status of the document. * If document is indexed successfully, the index_time field is populated. * Otherwise, if document is not indexed due to errors, t…
      - **indexTime** `string` — The time when the document was indexed. If this field is populated, it means the document has been indexed. While documents typically become searchable within seconds of indexin…
      - **errorSamples** `array<GoogleRpcStatus>` — A sample of errors encountered while indexing the document. If this field is populated, the document is not indexed due to errors.
      - **pendingMessage** `string` — Immutable. The message indicates the document index is in progress. If this field is populated, the document index is pending.
    - **parentDocumentId** `string` — The identifier of the parent document. Currently supports at most two level document hierarchy. Id should conform to [RFC-1034](https://tools.ietf.org/html/rfc1034) standard wit…
    - **id** `string` — Immutable. The identifier of the document. Id should conform to [RFC-1034](https://tools.ietf.org/html/rfc1034) standard with a length limit of 128 characters.
    - **schemaId** `string` — The identifier of the schema located in the same data store.
    - **content** `DocumentContent` — The unstructured data linked to this document. Content can only be set and must be set if this document is under a `CONTENT_REQUIRED` data store.
      - **rawBytes** `string` — The content represented as a stream of bytes. The maximum length is 1,000,000 bytes (1 MB / ~0.95 MiB). Note: As with all `bytes` fields, this field is represented as pure binar…
      - **mimeType** `string` — The MIME type of the content. Supported types: * `application/pdf` (PDF, only native PDFs are supported for now) * `text/html` (HTML) * `text/plain` (TXT) * `application/xml` or…
      - **uri** `string` — The URI of the content. Only Cloud Storage URIs (e.g. `gs://bucket-name/path/to/file`) are supported. The maximum file size is 2.5 MB for text-based formats, 200 MB for other fo…
    - **structData** `object` — The structured JSON data for the document. It should conform to the registered Schema or an `INVALID_ARGUMENT` error is thrown.
  - **score** `number` — The score of each suggestion. The score is in the range of [0, 1].
  - **destinationUri** `string` — The destination uri of the person suggestion.
  - **personType** `enum` — enum: PERSON_TYPE_UNSPECIFIED, CLOUD_IDENTITY, THIRD_PARTY_IDENTITY — The type of the person.
  - **suggestion** `string` — The suggestion for the query.
  - **dataStore** `string` — The name of the dataStore that this suggestion belongs to.
  - **displayPhotoUri** `string` — The photo uri of the person suggestion.

