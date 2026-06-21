# rank — semantic reranking

Rerank a candidate list of records by relevance to a query (e.g. ordering review findings or entity matches).

## rankingConfigs.rank


Rerank records against a query.

- **HTTP**: `POST` `https://discoveryengine.googleapis.com/v1alpha/projects/{projectsId}/locations/{locationsId}/rankingConfigs/{rankingConfigsId}:rank`
- **Method id**: `projects.locations.rankingConfigs.rank`
- **Scopes**: `https://www.googleapis.com/auth/cloud-platform, https://www.googleapis.com/auth/discoveryengine.readwrite, https://www.googleapis.com/auth/discoveryengine.serving.readwrite`

### Request body — `RankRequest`

- **model** `string` — The identifier of the model to use. It is one of: * `semantic-ranker-512@latest`: Semantic ranking model with maximum input token size 512. It is set to `semantic-ranker-512@lat…
- **topN** `integer` — The number of results to return. If this is unset or no bigger than zero, returns all results.
- **userLabels** `object` — The user labels applied to a resource must meet the following requirements: * Each resource can have multiple labels, up to a maximum of 64. * Each label must be a key-value pai…
- **query** `string` — The query to use.
- **ignoreRecordDetailsInResponse** `boolean` — If true, the response will contain only record ID and score. By default, it is false, the response will contain record details.
- **records** `array<RankingRecord>` — Required. A list of records to rank.
  - **score** `number` — The score of this record based on the given query and selected model. The score will be rounded to 4 decimal places. If the score is close to 0, it will be rounded to 0.00001 to…
  - **title** `string` — The title of the record. Empty by default. At least one of title or content should be set otherwise an INVALID_ARGUMENT error is thrown.
  - **content** `string` — The content of the record. Empty by default. At least one of title or content should be set otherwise an INVALID_ARGUMENT error is thrown.
  - **id** `string` — The unique ID to represent the record.

### Response — `RankResponse`

- **records** `array<RankingRecord>` — A list of records sorted by descending score.
  - **score** `number` — The score of this record based on the given query and selected model. The score will be rounded to 4 decimal places. If the score is close to 0, it will be rounded to 0.00001 to…
  - **title** `string` — The title of the record. Empty by default. At least one of title or content should be set otherwise an INVALID_ARGUMENT error is thrown.
  - **content** `string` — The content of the record. Empty by default. At least one of title or content should be set otherwise an INVALID_ARGUMENT error is thrown.
  - **id** `string` — The unique ID to represent the record.

