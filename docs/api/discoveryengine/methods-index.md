# Discovery Engine API (v1alpha) — full method index

Source: `https://discoveryengine.googleapis.com/$discovery/rest?version=v1alpha` (fetched 2026-06-21). 365 methods total.

Detailed request/response references for the assistant-relevant verbs are in the sibling files (`streamAssist.md`, `search.md`, `answer.md`, `sessions.md`, `autocomplete.md`, `grounding-check.md`, `ranking.md`, `assistant-config.md`). This index lists every method.

## `addPatientFilter` (1)

- `POST  ` `projects.locations.dataStores.addPatientFilter` — Adds a group of patient IDs as a patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. The filter group must be a FHIR resource name of type Group, and the filter will be constructed from the direct members of the group which are Patient resources.

## `audioOverviews` (2)

- `DELETE` `projects.locations.notebooks.audioOverviews.delete` — Deletes an audio overview.
- `POST  ` `projects.locations.notebooks.audioOverviews.create` — Generates a new audio overview.

## `batchDelete` (1)

- `POST  ` `projects.locations.notebooks.batchDelete` — Batch deletes Notebooks.

## `batchUpdateUserLicenses` (1)

- `POST  ` `projects.locations.userStores.batchUpdateUserLicenses` — Updates the User License. This method is used for batch assign/unassign licenses to users.

## `billingAccountLicenseConfigs` (4)

- `GET   ` `billingAccounts.billingAccountLicenseConfigs.get` — Gets a BillingAccountLicenseConfig.
- `GET   ` `billingAccounts.billingAccountLicenseConfigs.list` — Lists all BillingAccountLicenseConfigs for a given billing account.
- `POST  ` `billingAccounts.billingAccountLicenseConfigs.distributeLicenseConfig` — Distributes a LicenseConfig from billing account level to project level.
- `POST  ` `billingAccounts.billingAccountLicenseConfigs.retractLicenseConfig` — This method is called from the billing account side to retract the LicenseConfig from the given project back to the billing account.

## `branches` (16)

- `DELETE` `projects.locations.dataStores.branches.documents.delete` — Deletes a Document.
- `GET   ` `projects.locations.dataStores.branches.batchGetDocumentsMetadata` — Gets index freshness metadata for Documents. Supported for website search only.
- `GET   ` `projects.locations.dataStores.branches.documents.chunks.get` — Gets a Document.
- `GET   ` `projects.locations.dataStores.branches.documents.chunks.list` — Gets a list of Chunks.
- `GET   ` `projects.locations.dataStores.branches.documents.get` — Gets a Document.
- `GET   ` `projects.locations.dataStores.branches.documents.getProcessedDocument` — Gets the parsed layout information for a Document.
- `GET   ` `projects.locations.dataStores.branches.documents.list` — Gets a list of Documents.
- `GET   ` `projects.locations.dataStores.branches.get` — Retrieves a Branch.
- `GET   ` `projects.locations.dataStores.branches.list` — Lists all Branchs under the specified parent DataStore.
- `GET   ` `projects.locations.dataStores.branches.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.dataStores.branches.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `PATCH ` `projects.locations.dataStores.branches.documents.patch` — Updates a Document.
- `POST  ` `projects.locations.dataStores.branches.documents.create` — Creates a Document.
- `POST  ` `projects.locations.dataStores.branches.documents.import` — Bulk import of multiple Documents. Request processing may be synchronous. Non-existing items are created. Note: It is possible for a subset of the Documents to be successfully updated.
- `POST  ` `projects.locations.dataStores.branches.documents.purge` — Permanently deletes all selected Documents in a branch. This process is asynchronous. Depending on the number of Documents to be deleted, this operation can take hours to complete. Before the delete operation completes, some Documents might still be returned by DocumentService.GetDocument or DocumentService.ListDocuments. To get a list of the Documents to be deleted, set PurgeDocumentsRequest.force to false.
- `POST  ` `projects.locations.dataStores.branches.operations.cancel` — Starts asynchronous cancellation on a long-running operation. The server makes a best effort to cancel the operation, but success is not guaranteed. If the server doesn't support this method, it returns `google.rpc.Code.UNIMPLEMENTED`. Clients can use Operations.GetOperation or other methods to check whether the cancellation succeeded or whether the operation completed despite cancellation. On successful cancellation, the operation is not deleted; instead, it becomes an operation with an Operation.error value with a google.rpc.Status.code of `1`, corresponding to `Code.CANCELLED`.

## `check` (1)

- `POST  ` `projects.locations.groundingConfigs.check` — Performs a grounding check.

## `checkRequirement` (1)

- `POST  ` `projects.locations.requirements.checkRequirement` — Check a particular requirement.

## `collect` (1)

- `GET   ` `projects.locations.userEvents.collect` — Writes a single user event from the browser. This uses a GET request to due to browser restriction of POST-ing to a third-party domain. This method is used only by the Discovery Engine API JavaScript pixel and Google Tag Manager. Users should not call this method directly.

## `completeQuery` (1)

- `GET   ` `projects.locations.dataStores.completeQuery` — Completes the specified user input with keyword suggestions.

## `completionConfig` (1)

- `POST  ` `projects.locations.dataStores.completionConfig.completeQuery` — Completes the user input with advanced keyword suggestions.

## `completionSuggestions` (2)

- `POST  ` `projects.locations.dataStores.completionSuggestions.import` — Imports CompletionSuggestions for a DataStore.
- `POST  ` `projects.locations.dataStores.completionSuggestions.purge` — Permanently deletes all CompletionSuggestions for a DataStore.

## `controls` (5)

- `DELETE` `projects.locations.dataStores.controls.delete` — Deletes a Control. If the Control to delete does not exist, a NOT_FOUND error is returned.
- `GET   ` `projects.locations.dataStores.controls.get` — Gets a Control.
- `GET   ` `projects.locations.dataStores.controls.list` — Lists all Controls by their parent DataStore.
- `PATCH ` `projects.locations.dataStores.controls.patch` — Updates a Control. Control action type cannot be changed. If the Control to update does not exist, a NOT_FOUND error is returned.
- `POST  ` `projects.locations.dataStores.controls.create` — Creates a Control. By default 1000 controls are allowed for a data store. A request can be submitted to adjust this limit. If the Control to create already exists, an ALREADY_EXISTS error is returned.

## `conversations` (6)

- `DELETE` `projects.locations.dataStores.conversations.delete` — Deletes a Conversation. If the Conversation to delete does not exist, a NOT_FOUND error is returned.
- `GET   ` `projects.locations.dataStores.conversations.get` — Gets a Conversation.
- `GET   ` `projects.locations.dataStores.conversations.list` — Lists all Conversations by their parent DataStore.
- `PATCH ` `projects.locations.dataStores.conversations.patch` — Updates a Conversation. Conversation action type cannot be changed. If the Conversation to update does not exist, a NOT_FOUND error is returned.
- `POST  ` `projects.locations.dataStores.conversations.converse` — Converses a conversation.
- `POST  ` `projects.locations.dataStores.conversations.create` — Creates a Conversation. If the Conversation to create already exists, an ALREADY_EXISTS error is returned.

## `create` (7)

- `POST  ` `projects.locations.authorizations.create` — Creates an Authorization.
- `POST  ` `projects.locations.dataStores.create` — Creates a DataStore. DataStore is for storing Documents. To serve these documents for Search, or Recommendation use case, an Engine needs to be created separately.
- `POST  ` `projects.locations.evaluations.create` — Creates a Evaluation. Upon creation, the evaluation will be automatically triggered and begin execution.
- `POST  ` `projects.locations.identityMappingStores.create` — Creates a new Identity Mapping Store.
- `POST  ` `projects.locations.licenseConfigs.create` — Creates a LicenseConfig This method should only be used for creating NotebookLm licenses or Gemini Enterprise free trial licenses.
- `POST  ` `projects.locations.notebooks.create` — Creates a notebook.
- `POST  ` `projects.locations.sampleQuerySets.create` — Creates a SampleQuerySet

## `dataConnector` (8)

- `DELETE` `projects.locations.collections.dataConnector.mcp` — ServeMcpDeleteRequest serves a MCP DELETE request.
- `GET   ` `projects.locations.collections.dataConnector.checkRefreshToken` — Deprecated: Checks the existence of a refresh token for the EUC user for a given connection and returns its details. Use AcquireAccessToken instead and then check the validity of the returned token by asking the 3rd party system. There's no way to know for sure if a refresh token is valid without asking the 3rd party system.
- `GET   ` `projects.locations.collections.dataConnector.connectorRuns.list` — Lists the ConnectorRuns of a DataConnector.
- `GET   ` `projects.locations.collections.dataConnector.getConnectorSecret` — Get the secret for the associated connector.
- `GET   ` `projects.locations.collections.dataConnector.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataConnector.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `POST  ` `projects.locations.collections.dataConnector.acquireAccessToken` — Uses the per-user refresh token minted with AcquireAndStoreRefreshToken to generate and return a new access token and its details. Takes the access token from cache if available. Rotates the stored refresh token if needed. Uses the end user identity to return the user specific access token. Does *not* return the credentials configured by the administrator. Used by action execution and UI.
- `POST  ` `projects.locations.collections.dataConnector.startConnectorRun` — Starts an immediate synchronization process for a DataConnector. Third Party Connector Users must specify which entities should be synced. FHIR Connectors must provide a timestamp to indicate the point in time from which data should be synced.

## `dataStores` (102)

- `DELETE` `projects.locations.collections.dataStores.branches.documents.delete` — Deletes a Document.
- `DELETE` `projects.locations.collections.dataStores.controls.delete` — Deletes a Control. If the Control to delete does not exist, a NOT_FOUND error is returned.
- `DELETE` `projects.locations.collections.dataStores.conversations.delete` — Deletes a Conversation. If the Conversation to delete does not exist, a NOT_FOUND error is returned.
- `DELETE` `projects.locations.collections.dataStores.delete` — Deletes a DataStore.
- `DELETE` `projects.locations.collections.dataStores.schemas.delete` — Deletes a Schema.
- `DELETE` `projects.locations.collections.dataStores.servingConfigs.delete` — Deletes a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `DELETE` `projects.locations.collections.dataStores.sessions.delete` — Deletes a Session. If the Session to delete does not exist, a NOT_FOUND error is returned.
- `DELETE` `projects.locations.collections.dataStores.siteSearchEngine.sitemaps.delete` — Deletes a Sitemap.
- `DELETE` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.delete` — Deletes a TargetSite.
- `GET   ` `projects.locations.collections.dataStores.branches.batchGetDocumentsMetadata` — Gets index freshness metadata for Documents. Supported for website search only.
- `GET   ` `projects.locations.collections.dataStores.branches.documents.chunks.get` — Gets a Document.
- `GET   ` `projects.locations.collections.dataStores.branches.documents.chunks.list` — Gets a list of Chunks.
- `GET   ` `projects.locations.collections.dataStores.branches.documents.get` — Gets a Document.
- `GET   ` `projects.locations.collections.dataStores.branches.documents.getProcessedDocument` — Gets the parsed layout information for a Document.
- `GET   ` `projects.locations.collections.dataStores.branches.documents.list` — Gets a list of Documents.
- `GET   ` `projects.locations.collections.dataStores.branches.get` — Retrieves a Branch.
- `GET   ` `projects.locations.collections.dataStores.branches.list` — Lists all Branchs under the specified parent DataStore.
- `GET   ` `projects.locations.collections.dataStores.branches.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.branches.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.completeQuery` — Completes the specified user input with keyword suggestions.
- `GET   ` `projects.locations.collections.dataStores.controls.get` — Gets a Control.
- `GET   ` `projects.locations.collections.dataStores.controls.list` — Lists all Controls by their parent DataStore.
- `GET   ` `projects.locations.collections.dataStores.conversations.get` — Gets a Conversation.
- `GET   ` `projects.locations.collections.dataStores.conversations.list` — Lists all Conversations by their parent DataStore.
- `GET   ` `projects.locations.collections.dataStores.customModels.list` — Gets a list of all the custom models.
- `GET   ` `projects.locations.collections.dataStores.get` — Gets a DataStore.
- `GET   ` `projects.locations.collections.dataStores.getCompletionConfig` — Gets a CompletionConfig
- `GET   ` `projects.locations.collections.dataStores.getDocumentProcessingConfig` — Gets a DocumentProcessingConfig.
- `GET   ` `projects.locations.collections.dataStores.getSiteSearchEngine` — Gets the SiteSearchEngine.
- `GET   ` `projects.locations.collections.dataStores.list` — Lists all the DataStores associated with the project.
- `GET   ` `projects.locations.collections.dataStores.models.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.models.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.schemas.get` — Gets a Schema.
- `GET   ` `projects.locations.collections.dataStores.schemas.list` — Gets a list of Schemas.
- `GET   ` `projects.locations.collections.dataStores.schemas.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.schemas.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.servingConfigs.get` — Gets a ServingConfig. Returns a NotFound error if the ServingConfig does not exist.
- `GET   ` `projects.locations.collections.dataStores.servingConfigs.list` — Lists all ServingConfigs linked to this dataStore.
- `GET   ` `projects.locations.collections.dataStores.sessions.answers.get` — Gets a Answer.
- `GET   ` `projects.locations.collections.dataStores.sessions.get` — Gets a Session.
- `GET   ` `projects.locations.collections.dataStores.sessions.list` — Lists all Sessions by their parent DataStore.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.fetchDomainVerificationStatus` — Returns list of target sites with its domain verification status. This method can only be called under data store with BASIC_SITE_SEARCH state at the moment.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.getUriPatternDocumentData` — Gets the URI Pattern to Document data mapping for an Advanced Site Search DataStore.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.sitemaps.fetch` — Fetch Sitemaps in a DataStore.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.get` — Gets a TargetSite.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.list` — Gets a list of TargetSites.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.dataStores.userEvents.collect` — Writes a single user event from the browser. This uses a GET request to due to browser restriction of POST-ing to a third-party domain. This method is used only by the Discovery Engine API JavaScript pixel and Google Tag Manager. Users should not call this method directly.
- `GET   ` `projects.locations.collections.dataStores.widgetConfigs.get` — Gets a WidgetConfig.
- `PATCH ` `projects.locations.collections.dataStores.branches.documents.patch` — Updates a Document.
- `PATCH ` `projects.locations.collections.dataStores.controls.patch` — Updates a Control. Control action type cannot be changed. If the Control to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.dataStores.conversations.patch` — Updates a Conversation. Conversation action type cannot be changed. If the Conversation to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.dataStores.patch` — Updates a DataStore
- `PATCH ` `projects.locations.collections.dataStores.schemas.patch` — Updates a Schema.
- `PATCH ` `projects.locations.collections.dataStores.servingConfigs.patch` — Updates a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `PATCH ` `projects.locations.collections.dataStores.sessions.patch` — Updates a Session. Session action type cannot be changed. If the Session to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.patch` — Updates a TargetSite.
- `PATCH ` `projects.locations.collections.dataStores.updateCompletionConfig` — Updates the CompletionConfigs.
- `PATCH ` `projects.locations.collections.dataStores.updateDocumentProcessingConfig` — Updates the DocumentProcessingConfig. DocumentProcessingConfig is a singleon resource of DataStore. It's empty when DataStore is created. The first call to this method will set up DocumentProcessingConfig.
- `PATCH ` `projects.locations.collections.dataStores.widgetConfigs.patch` — Update a WidgetConfig.
- `POST  ` `projects.locations.collections.dataStores.addPatientFilter` — Adds a group of patient IDs as a patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. The filter group must be a FHIR resource name of type Group, and the filter will be constructed from the direct members of the group which are Patient resources.
- `POST  ` `projects.locations.collections.dataStores.branches.documents.create` — Creates a Document.
- `POST  ` `projects.locations.collections.dataStores.branches.documents.import` — Bulk import of multiple Documents. Request processing may be synchronous. Non-existing items are created. Note: It is possible for a subset of the Documents to be successfully updated.
- `POST  ` `projects.locations.collections.dataStores.branches.documents.purge` — Permanently deletes all selected Documents in a branch. This process is asynchronous. Depending on the number of Documents to be deleted, this operation can take hours to complete. Before the delete operation completes, some Documents might still be returned by DocumentService.GetDocument or DocumentService.ListDocuments. To get a list of the Documents to be deleted, set PurgeDocumentsRequest.force to false.
- `POST  ` `projects.locations.collections.dataStores.branches.operations.cancel` — Starts asynchronous cancellation on a long-running operation. The server makes a best effort to cancel the operation, but success is not guaranteed. If the server doesn't support this method, it returns `google.rpc.Code.UNIMPLEMENTED`. Clients can use Operations.GetOperation or other methods to check whether the cancellation succeeded or whether the operation completed despite cancellation. On successful cancellation, the operation is not deleted; instead, it becomes an operation with an Operation.error value with a google.rpc.Status.code of `1`, corresponding to `Code.CANCELLED`.
- `POST  ` `projects.locations.collections.dataStores.completionConfig.completeQuery` — Completes the user input with advanced keyword suggestions.
- `POST  ` `projects.locations.collections.dataStores.completionSuggestions.import` — Imports CompletionSuggestions for a DataStore.
- `POST  ` `projects.locations.collections.dataStores.completionSuggestions.purge` — Permanently deletes all CompletionSuggestions for a DataStore.
- `POST  ` `projects.locations.collections.dataStores.controls.create` — Creates a Control. By default 1000 controls are allowed for a data store. A request can be submitted to adjust this limit. If the Control to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.dataStores.conversations.converse` — Converses a conversation.
- `POST  ` `projects.locations.collections.dataStores.conversations.create` — Creates a Conversation. If the Conversation to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.dataStores.create` — Creates a DataStore. DataStore is for storing Documents. To serve these documents for Search, or Recommendation use case, an Engine needs to be created separately.
- `POST  ` `projects.locations.collections.dataStores.deletePatientFilter` — Deletes the entire patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter.
- `POST  ` `projects.locations.collections.dataStores.removePatientFilter` — Removes a group of patient IDs from the patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter. The filter group must be a FHIR resource name of type Group, and the list of patient IDs to remove will be constructed from the direct members of the group which are Patient resources.
- `POST  ` `projects.locations.collections.dataStores.replacePatientFilter` — Replaces the patient filter for the data store. This method is essentially a combination of DeletePatientFilters and AddPatientFilter. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter. The filter group must be a FHIR resource name of type Group, and the new filter will be constructed from the direct members of the group which are Patient resources.
- `POST  ` `projects.locations.collections.dataStores.schemas.create` — Creates a Schema.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.answer` — Answer query method.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.create` — Creates a ServingConfig. Note: The Google Cloud console works only with the default serving config. Additional ServingConfigs can be created and managed only via the API. A maximum of 100 ServingConfigs are allowed in an Engine, otherwise a RESOURCE_EXHAUSTED error is returned.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.recommend` — Makes a recommendation, which requires a contextual user event.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.search` — Performs a search.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.searchLite` — Performs a search. Similar to the SearchService.Search method, but a lite version that allows API key for authentication, where OAuth and IAM checks are not required. Only public website search is supported by this method. If data stores and engines not associated with public website search are specified, a `FAILED_PRECONDITION` error is returned. This method can be used for easy onboarding without having to implement an authentication backend. However, it is strongly recommended to use SearchService.Search instead with required OAuth and IAM checks to provide better data security.
- `POST  ` `projects.locations.collections.dataStores.servingConfigs.streamAnswer` — Answer query method (streaming). It takes one AnswerQueryRequest and returns multiple AnswerQueryResponse messages in a stream.
- `POST  ` `projects.locations.collections.dataStores.sessions.create` — Creates a Session. If the Session to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.batchVerifyTargetSites` — Verify target sites' ownership and validity. This API sends all the target sites under site search engine for verification.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.disableAdvancedSiteSearch` — Downgrade from advanced site search to basic site search.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.enableAdvancedSiteSearch` — Upgrade from basic site search to advanced site search.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.recrawlUris` — Request on-demand recrawl for a list of URIs.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.setUriPatternDocumentData` — Sets the URI Pattern to Document data mapping for an Advanced Site Search DataStore.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.sitemaps.create` — Creates a Sitemap.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.batchCreate` — Creates TargetSite in a batch.
- `POST  ` `projects.locations.collections.dataStores.siteSearchEngine.targetSites.create` — Creates a TargetSite.
- `POST  ` `projects.locations.collections.dataStores.suggestionDenyListEntries.import` — Imports all SuggestionDenyListEntry for a DataStore.
- `POST  ` `projects.locations.collections.dataStores.suggestionDenyListEntries.purge` — Permanently deletes all SuggestionDenyListEntry for a DataStore.
- `POST  ` `projects.locations.collections.dataStores.trainCustomModel` — Trains a custom model.
- `POST  ` `projects.locations.collections.dataStores.userEvents.import` — Bulk import of user events. Request processing might be synchronous. Events that already exist are skipped. Use this method for backfilling historical user events. Operation.response is of type ImportResponse. Note that it is possible for a subset of the items to be successfully inserted. Operation.metadata is of type ImportMetadata.
- `POST  ` `projects.locations.collections.dataStores.userEvents.purge` — Deletes permanently all user events specified by the filter provided. Depending on the number of events specified by the filter, this operation could take hours or days to complete. To test a filter, use the list command first.
- `POST  ` `projects.locations.collections.dataStores.userEvents.write` — Writes a single user event.

## `delete` (6)

- `DELETE` `projects.locations.authorizations.delete` — Deletes an Authorization.
- `DELETE` `projects.locations.cmekConfigs.delete` — De-provisions a CmekConfig.
- `DELETE` `projects.locations.collections.delete` — Deletes a Collection.
- `DELETE` `projects.locations.dataStores.delete` — Deletes a DataStore.
- `DELETE` `projects.locations.identityMappingStores.delete` — Deletes the Identity Mapping Store.
- `DELETE` `projects.locations.sampleQuerySets.delete` — Deletes a SampleQuerySet.

## `deletePatientFilter` (1)

- `POST  ` `projects.locations.dataStores.deletePatientFilter` — Deletes the entire patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter.

## `engines` (70)

- `DELETE` `projects.locations.collections.engines.assistants.agents.delete` — Deletes an Agent.
- `DELETE` `projects.locations.collections.engines.assistants.cannedQueries.delete` — Deletes a CannedQuery.
- `DELETE` `projects.locations.collections.engines.assistants.delete` — Deletes an Assistant.
- `DELETE` `projects.locations.collections.engines.controls.delete` — Deletes a Control. If the Control to delete does not exist, a NOT_FOUND error is returned.
- `DELETE` `projects.locations.collections.engines.conversations.delete` — Deletes a Conversation. If the Conversation to delete does not exist, a NOT_FOUND error is returned.
- `DELETE` `projects.locations.collections.engines.delete` — Deletes an Engine.
- `DELETE` `projects.locations.collections.engines.servingConfigs.delete` — Deletes a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `DELETE` `projects.locations.collections.engines.sessions.delete` — Deletes a Session. If the Session to delete does not exist, a NOT_FOUND error is returned.
- `GET   ` `projects.locations.collections.engines.analytics.getConfig` — Gets the AnalyticsConfig.
- `GET   ` `projects.locations.collections.engines.assistants.agents.get` — Gets an Agent.
- `GET   ` `projects.locations.collections.engines.assistants.agents.list` — Lists all Agents under an Assistant which were created by the caller.
- `GET   ` `projects.locations.collections.engines.assistants.agents.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.engines.assistants.cannedQueries.get` — Gets a CannedQuery.
- `GET   ` `projects.locations.collections.engines.assistants.cannedQueries.list` — Lists all CannedQuerys under an Assistant.
- `GET   ` `projects.locations.collections.engines.assistants.get` — Gets an Assistant.
- `GET   ` `projects.locations.collections.engines.assistants.list` — Lists all Assistants under an Engine.
- `GET   ` `projects.locations.collections.engines.controls.get` — Gets a Control.
- `GET   ` `projects.locations.collections.engines.controls.list` — Lists all Controls by their parent DataStore.
- `GET   ` `projects.locations.collections.engines.conversations.get` — Gets a Conversation.
- `GET   ` `projects.locations.collections.engines.conversations.list` — Lists all Conversations by their parent DataStore.
- `GET   ` `projects.locations.collections.engines.get` — Gets an Engine.
- `GET   ` `projects.locations.collections.engines.getCompletionConfig` — Gets a CompletionConfig
- `GET   ` `projects.locations.collections.engines.getIamPolicy` — Gets the IAM access control policy for an Engine. A `NOT_FOUND` error is returned if the resource does not exist. An empty policy is returned if the resource exists but does not have a policy set on it.
- `GET   ` `projects.locations.collections.engines.getWorkspaceSettings` — Get Workspace settings for the end user.
- `GET   ` `projects.locations.collections.engines.list` — Lists all the Engines associated with the project.
- `GET   ` `projects.locations.collections.engines.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.engines.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.collections.engines.servingConfigs.get` — Gets a ServingConfig. Returns a NotFound error if the ServingConfig does not exist.
- `GET   ` `projects.locations.collections.engines.servingConfigs.list` — Lists all ServingConfigs linked to this dataStore.
- `GET   ` `projects.locations.collections.engines.sessions.alphaEvolveExperiments.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.engines.sessions.answers.get` — Gets a Answer.
- `GET   ` `projects.locations.collections.engines.sessions.files.list` — Lists metadata for all files in the current session.
- `GET   ` `projects.locations.collections.engines.sessions.get` — Gets a Session.
- `GET   ` `projects.locations.collections.engines.sessions.list` — Lists all Sessions by their parent DataStore.
- `GET   ` `projects.locations.collections.engines.sessions.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.engines.widgetConfigs.get` — Gets a WidgetConfig.
- `PATCH ` `projects.locations.collections.engines.analytics.updateConfig` — Updates the AnalyticsConfig for analytics.
- `PATCH ` `projects.locations.collections.engines.assistants.agents.patch` — Updates an Agent
- `PATCH ` `projects.locations.collections.engines.assistants.cannedQueries.patch` — Updates a CannedQuery.
- `PATCH ` `projects.locations.collections.engines.assistants.patch` — Updates an Assistant
- `PATCH ` `projects.locations.collections.engines.controls.patch` — Updates a Control. Control action type cannot be changed. If the Control to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.engines.conversations.patch` — Updates a Conversation. Conversation action type cannot be changed. If the Conversation to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.engines.patch` — Updates an Engine
- `PATCH ` `projects.locations.collections.engines.servingConfigs.patch` — Updates a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `PATCH ` `projects.locations.collections.engines.sessions.patch` — Updates a Session. Session action type cannot be changed. If the Session to update does not exist, a NOT_FOUND error is returned.
- `PATCH ` `projects.locations.collections.engines.updateCompletionConfig` — Updates the CompletionConfigs.
- `PATCH ` `projects.locations.collections.engines.widgetConfigs.patch` — Update a WidgetConfig.
- `POST  ` `projects.locations.collections.engines.analytics.exportMetrics` — Exports metrics.
- `POST  ` `projects.locations.collections.engines.assistants.agents.create` — Creates an Agent.
- `POST  ` `projects.locations.collections.engines.assistants.agents.files.import` — Imports a file to an Agent. Currently only No-Code agents are supported.
- `POST  ` `projects.locations.collections.engines.assistants.cannedQueries.create` — Creates a CannedQuery.
- `POST  ` `projects.locations.collections.engines.assistants.create` — Creates an Assistant.
- `POST  ` `projects.locations.collections.engines.assistants.streamAssist` — Assists the user with a query in a streaming fashion.
- `POST  ` `projects.locations.collections.engines.completionConfig.completeQuery` — Completes the user input with advanced keyword suggestions.
- `POST  ` `projects.locations.collections.engines.completionConfig.removeSuggestion` — Removes the search history suggestion in an engine for a user. This will remove the suggestion from being returned in the AdvancedCompleteQueryResponse.recent_search_suggestions for this user. If the user searches the same suggestion again, the new history will override and suggest this suggestion again.
- `POST  ` `projects.locations.collections.engines.controls.create` — Creates a Control. By default 1000 controls are allowed for a data store. A request can be submitted to adjust this limit. If the Control to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.engines.conversations.converse` — Converses a conversation.
- `POST  ` `projects.locations.collections.engines.conversations.create` — Creates a Conversation. If the Conversation to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.engines.create` — Creates an Engine.
- `POST  ` `projects.locations.collections.engines.pause` — Pauses the training of an existing Engine. Only applicable if SolutionType is SOLUTION_TYPE_RECOMMENDATION.
- `POST  ` `projects.locations.collections.engines.resume` — Resumes the training of an existing Engine. Only applicable if SolutionType is SOLUTION_TYPE_RECOMMENDATION.
- `POST  ` `projects.locations.collections.engines.servingConfigs.answer` — Answer query method.
- `POST  ` `projects.locations.collections.engines.servingConfigs.create` — Creates a ServingConfig. Note: The Google Cloud console works only with the default serving config. Additional ServingConfigs can be created and managed only via the API. A maximum of 100 ServingConfigs are allowed in an Engine, otherwise a RESOURCE_EXHAUSTED error is returned.
- `POST  ` `projects.locations.collections.engines.servingConfigs.recommend` — Makes a recommendation, which requires a contextual user event.
- `POST  ` `projects.locations.collections.engines.servingConfigs.search` — Performs a search.
- `POST  ` `projects.locations.collections.engines.servingConfigs.searchLite` — Performs a search. Similar to the SearchService.Search method, but a lite version that allows API key for authentication, where OAuth and IAM checks are not required. Only public website search is supported by this method. If data stores and engines not associated with public website search are specified, a `FAILED_PRECONDITION` error is returned. This method can be used for easy onboarding without having to implement an authentication backend. However, it is strongly recommended to use SearchService.Search instead with required OAuth and IAM checks to provide better data security.
- `POST  ` `projects.locations.collections.engines.servingConfigs.streamAnswer` — Answer query method (streaming). It takes one AnswerQueryRequest and returns multiple AnswerQueryResponse messages in a stream.
- `POST  ` `projects.locations.collections.engines.sessions.create` — Creates a Session. If the Session to create already exists, an ALREADY_EXISTS error is returned.
- `POST  ` `projects.locations.collections.engines.setIamPolicy` — Sets the IAM access control policy for an Engine. A `NOT_FOUND` error is returned if the resource does not exist. **Important:** When setting a policy directly on an Engine resource, the only recommended roles in the bindings are: `roles/discoveryengine.admin`, `roles/discoveryengine.agentspaceAdmin`, `roles/discoveryengine.user`, `roles/discoveryengine.agentspaceUser`, `roles/discoveryengine.viewer`, `roles/discoveryengine.agentspaceViewer`. Attempting to grant any other role will result in a warning in logging.
- `POST  ` `projects.locations.collections.engines.tune` — Tunes an existing Engine. Only applicable if SolutionType is SOLUTION_TYPE_RECOMMENDATION.

## `get` (11)

- `GET   ` `projects.locations.authorizations.get` — Gets an Authorization.
- `GET   ` `projects.locations.cmekConfigs.get` — Gets the CmekConfig.
- `GET   ` `projects.locations.collections.get` — Gets a Collection.
- `GET   ` `projects.locations.dataStores.get` — Gets a DataStore.
- `GET   ` `projects.locations.evaluations.get` — Gets a Evaluation.
- `GET   ` `projects.locations.identityMappingStores.get` — Gets the Identity Mapping Store.
- `GET   ` `projects.locations.licenseConfigs.get` — Gets a LicenseConfig.
- `GET   ` `projects.locations.notebooks.get` — Gets a notebook.
- `GET   ` `projects.locations.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.sampleQuerySets.get` — Gets a SampleQuerySet.
- `GET   ` `projects.locations.userStores.get` — Gets the User Store.

## `getCompletionConfig` (1)

- `GET   ` `projects.locations.dataStores.getCompletionConfig` — Gets a CompletionConfig

## `getDataConnector` (1)

- `GET   ` `projects.locations.collections.getDataConnector` — Gets the DataConnector. DataConnector is a singleton resource for each Collection.

## `getDocumentProcessingConfig` (1)

- `GET   ` `projects.locations.dataStores.getDocumentProcessingConfig` — Gets a DocumentProcessingConfig.

## `getSiteSearchEngine` (1)

- `GET   ` `projects.locations.dataStores.getSiteSearchEngine` — Gets the SiteSearchEngine.

## `import` (1)

- `POST  ` `projects.locations.userEvents.import` — Bulk import of user events. Request processing might be synchronous. Events that already exist are skipped. Use this method for backfilling historical user events. Operation.response is of type ImportResponse. Note that it is possible for a subset of the items to be successfully inserted. Operation.metadata is of type ImportMetadata.

## `importIdentityMappings` (1)

- `POST  ` `projects.locations.identityMappingStores.importIdentityMappings` — Imports a list of Identity Mapping Entries to an Identity Mapping Store.

## `licenseConfigsUsageStats` (1)

- `GET   ` `projects.locations.userStores.licenseConfigsUsageStats.list` — Lists all the LicenseConfigUsageStatss associated with the project.

## `list` (9)

- `GET   ` `projects.locations.authorizations.list` — Lists all Authorizations under an Engine.
- `GET   ` `projects.locations.cmekConfigs.list` — Lists all the CmekConfigs with the project.
- `GET   ` `projects.locations.collections.list` — Gets a list of Collections.
- `GET   ` `projects.locations.dataStores.list` — Lists all the DataStores associated with the project.
- `GET   ` `projects.locations.evaluations.list` — Gets a list of Evaluations.
- `GET   ` `projects.locations.identityMappingStores.list` — Lists all Identity Mapping Stores.
- `GET   ` `projects.locations.licenseConfigs.list` — Lists all the LicenseConfigs associated with the project.
- `GET   ` `projects.locations.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.sampleQuerySets.list` — Gets a list of SampleQuerySets.

## `listIdentityMappings` (1)

- `GET   ` `projects.locations.identityMappingStores.listIdentityMappings` — Lists Identity Mappings in an Identity Mapping Store.

## `listRecentlyViewed` (1)

- `GET   ` `projects.locations.notebooks.listRecentlyViewed` — Lists the notebooks ordered by last view time.

## `listResults` (1)

- `GET   ` `projects.locations.evaluations.listResults` — Gets a list of results for a given a Evaluation.

## `locations` (12)

- `GET   ` `projects.locations.completeExternalIdentities` — This method provides suggestions for users and groups managed in an external identity provider, based on the provided prefix.
- `GET   ` `projects.locations.getAclConfig` — Gets the AclConfig.
- `GET   ` `projects.locations.getCmekConfig` — Gets the CmekConfig.
- `GET   ` `projects.locations.queryConfigurablePricingUsageStats` — Queries configurable pricing usage stats for a project.
- `PATCH ` `projects.locations.updateAclConfig` — Default ACL configuration for use in a location of a customer's project. Updates will only reflect to new data stores. Existing data stores will still use the old value.
- `PATCH ` `projects.locations.updateCmekConfig` — Provisions a CMEK key for use in a location of a customer's project. This method will also conduct location validation on the provided cmekConfig to make sure the key is valid and can be used in the selected location.
- `POST  ` `projects.locations.estimateDataSize` — Estimates the data size to be used by a customer.
- `POST  ` `projects.locations.obtainCrawlRate` — Obtains the time series data of organic or dedicated crawl rate for monitoring. When dedicated crawl rate is not set, it will return vertex AI's organic crawl rate time series. Organic crawl means Google automatically crawl the internet at its own convenience. When dedicated crawl rate is set, it will return vertex AI's dedicated crawl rate time series.
- `POST  ` `projects.locations.removeDedicatedCrawlRate` — Removes the dedicated crawl rate for a craw_rate_scope. If the dedicated crawl rate was set, this will disable vertex AI's crawl bot from using the dedicated crawl rate for crawling. If the dedicated crawl rate was not set, this is a no-op.
- `POST  ` `projects.locations.setDedicatedCrawlRate` — Sets the dedicated crawl rate for a crawl_rate_scope. If the dedicated crawl rate was not set, this will enable vertex AI's crawl bot to use the new dedicated crawl rate for crawling. If the dedicated crawl rate was set, vertex AI's crawl bot will try to update the rate to the new value. If the new value is too high, the crawl bot may crawl at a lower rate to avoid overloading the user's website.
- `POST  ` `projects.locations.setUpDataConnector` — Creates a Collection and sets up the DataConnector for it. To stop a DataConnector after setup, use the CollectionService.DeleteCollection method.
- `POST  ` `projects.locations.setUpDataConnectorV2` — Creates a Collection and sets up the DataConnector for it. To stop a DataConnector after setup, use the CollectionService.DeleteCollection method.

## `models` (2)

- `GET   ` `projects.locations.dataStores.models.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.dataStores.models.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.

## `operations` (13)

- `GET   ` `projects.locations.collections.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.collections.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.dataStores.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.dataStores.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.evaluations.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.identityMappingStores.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.identityMappingStores.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.locations.podcasts.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.sampleQuerySets.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.userStores.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.locations.userStores.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.
- `GET   ` `projects.operations.get` — Gets the latest state of a long-running operation. Clients can use this method to poll the operation result at intervals as recommended by the API service.
- `GET   ` `projects.operations.list` — Lists operations that match the specified filter in the request. If the server doesn't support this method, it returns `UNIMPLEMENTED`.

## `patch` (7)

- `PATCH ` `projects.locations.authorizations.patch` — Updates an Authorization
- `PATCH ` `projects.locations.cmekConfigs.patch` — Provisions a CMEK key for use in a location of a customer's project. This method will also conduct location validation on the provided cmekConfig to make sure the key is valid and can be used in the selected location.
- `PATCH ` `projects.locations.collections.patch` — Updates a Collection.
- `PATCH ` `projects.locations.dataStores.patch` — Updates a DataStore
- `PATCH ` `projects.locations.licenseConfigs.patch` — Updates the LicenseConfig
- `PATCH ` `projects.locations.sampleQuerySets.patch` — Updates a SampleQuerySet.
- `PATCH ` `projects.locations.userStores.patch` — Updates the User Store.

## `projects` (4)

- `GET   ` `projects.get` — Gets a Project. Returns NOT_FOUND when the project is not yet created.
- `PATCH ` `projects.patch` — Updates the editable settings of a Discovery Engine Project.
- `POST  ` `projects.provision` — Provisions the project resource. During the process, related systems will get prepared and initialized. Caller must read the [Terms for data use](https://cloud.google.com/retail/data-use-terms), and optionally specify in request to provide consent to that service terms.
- `POST  ` `projects.reportConsentChange` — Updates service terms for this project. This method can be used to retroactively accept the latest terms. Terms available for update: * [Terms for data use](https://cloud.google.com/retail/data-use-terms)

## `purgeIdentityMappings` (1)

- `POST  ` `projects.locations.identityMappingStores.purgeIdentityMappings` — Purges specified or all Identity Mapping Entries from an Identity Mapping Store.

## `rank` (1)

- `POST  ` `projects.locations.rankingConfigs.rank` — Ranks a list of text records based on the given input query.

## `removePatientFilter` (1)

- `POST  ` `projects.locations.dataStores.removePatientFilter` — Removes a group of patient IDs from the patient filter for the data store. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter. The filter group must be a FHIR resource name of type Group, and the list of patient IDs to remove will be constructed from the direct members of the group which are Patient resources.

## `replacePatientFilter` (1)

- `POST  ` `projects.locations.dataStores.replacePatientFilter` — Replaces the patient filter for the data store. This method is essentially a combination of DeletePatientFilters and AddPatientFilter. Patient filters are empty by default when a data store is created, and are stored in a separate table. The data store must first be created, and must be a healthcare data store. This method will fail if the data store does not have a patient filter. The filter group must be a FHIR resource name of type Group, and the new filter will be constructed from the direct members of the group which are Patient resources.

## `sampleQueries` (6)

- `DELETE` `projects.locations.sampleQuerySets.sampleQueries.delete` — Deletes a SampleQuery.
- `GET   ` `projects.locations.sampleQuerySets.sampleQueries.get` — Gets a SampleQuery.
- `GET   ` `projects.locations.sampleQuerySets.sampleQueries.list` — Gets a list of SampleQuerys.
- `PATCH ` `projects.locations.sampleQuerySets.sampleQueries.patch` — Updates a SampleQuery.
- `POST  ` `projects.locations.sampleQuerySets.sampleQueries.create` — Creates a SampleQuery
- `POST  ` `projects.locations.sampleQuerySets.sampleQueries.import` — Bulk import of multiple SampleQuerys. Sample queries that already exist may be deleted. Note: It is possible for a subset of the SampleQuerys to be successfully imported.

## `schemas` (5)

- `DELETE` `projects.locations.dataStores.schemas.delete` — Deletes a Schema.
- `GET   ` `projects.locations.dataStores.schemas.get` — Gets a Schema.
- `GET   ` `projects.locations.dataStores.schemas.list` — Gets a list of Schemas.
- `PATCH ` `projects.locations.dataStores.schemas.patch` — Updates a Schema.
- `POST  ` `projects.locations.dataStores.schemas.create` — Creates a Schema.

## `servingConfigs` (10)

- `DELETE` `projects.locations.dataStores.servingConfigs.delete` — Deletes a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `GET   ` `projects.locations.dataStores.servingConfigs.get` — Gets a ServingConfig. Returns a NotFound error if the ServingConfig does not exist.
- `GET   ` `projects.locations.dataStores.servingConfigs.list` — Lists all ServingConfigs linked to this dataStore.
- `PATCH ` `projects.locations.dataStores.servingConfigs.patch` — Updates a ServingConfig. Returns a NOT_FOUND error if the ServingConfig does not exist.
- `POST  ` `projects.locations.dataStores.servingConfigs.answer` — Answer query method.
- `POST  ` `projects.locations.dataStores.servingConfigs.create` — Creates a ServingConfig. Note: The Google Cloud console works only with the default serving config. Additional ServingConfigs can be created and managed only via the API. A maximum of 100 ServingConfigs are allowed in an Engine, otherwise a RESOURCE_EXHAUSTED error is returned.
- `POST  ` `projects.locations.dataStores.servingConfigs.recommend` — Makes a recommendation, which requires a contextual user event.
- `POST  ` `projects.locations.dataStores.servingConfigs.search` — Performs a search.
- `POST  ` `projects.locations.dataStores.servingConfigs.searchLite` — Performs a search. Similar to the SearchService.Search method, but a lite version that allows API key for authentication, where OAuth and IAM checks are not required. Only public website search is supported by this method. If data stores and engines not associated with public website search are specified, a `FAILED_PRECONDITION` error is returned. This method can be used for easy onboarding without having to implement an authentication backend. However, it is strongly recommended to use SearchService.Search instead with required OAuth and IAM checks to provide better data security.
- `POST  ` `projects.locations.dataStores.servingConfigs.streamAnswer` — Answer query method (streaming). It takes one AnswerQueryRequest and returns multiple AnswerQueryResponse messages in a stream.

## `sessions` (6)

- `DELETE` `projects.locations.dataStores.sessions.delete` — Deletes a Session. If the Session to delete does not exist, a NOT_FOUND error is returned.
- `GET   ` `projects.locations.dataStores.sessions.answers.get` — Gets a Answer.
- `GET   ` `projects.locations.dataStores.sessions.get` — Gets a Session.
- `GET   ` `projects.locations.dataStores.sessions.list` — Lists all Sessions by their parent DataStore.
- `PATCH ` `projects.locations.dataStores.sessions.patch` — Updates a Session. Session action type cannot be changed. If the Session to update does not exist, a NOT_FOUND error is returned.
- `POST  ` `projects.locations.dataStores.sessions.create` — Creates a Session. If the Session to create already exists, an ALREADY_EXISTS error is returned.

## `share` (1)

- `POST  ` `projects.locations.notebooks.share` — Shares a notebook to other accounts.

## `siteSearchEngine` (12)

- `DELETE` `projects.locations.dataStores.siteSearchEngine.sitemaps.delete` — Deletes a Sitemap.
- `DELETE` `projects.locations.dataStores.siteSearchEngine.targetSites.delete` — Deletes a TargetSite.
- `GET   ` `projects.locations.dataStores.siteSearchEngine.sitemaps.fetch` — Fetch Sitemaps in a DataStore.
- `GET   ` `projects.locations.dataStores.siteSearchEngine.targetSites.get` — Gets a TargetSite.
- `GET   ` `projects.locations.dataStores.siteSearchEngine.targetSites.list` — Gets a list of TargetSites.
- `PATCH ` `projects.locations.dataStores.siteSearchEngine.targetSites.patch` — Updates a TargetSite.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.disableAdvancedSiteSearch` — Downgrade from advanced site search to basic site search.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.enableAdvancedSiteSearch` — Upgrade from basic site search to advanced site search.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.recrawlUris` — Request on-demand recrawl for a list of URIs.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.sitemaps.create` — Creates a Sitemap.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.targetSites.batchCreate` — Creates TargetSite in a batch.
- `POST  ` `projects.locations.dataStores.siteSearchEngine.targetSites.create` — Creates a TargetSite.

## `sources` (3)

- `GET   ` `projects.locations.notebooks.sources.get` — Gets a Source.
- `POST  ` `projects.locations.notebooks.sources.batchCreate` — Creates a list of Sources.
- `POST  ` `projects.locations.notebooks.sources.batchDelete` — Deletes multiple sources

## `suggestionDenyListEntries` (2)

- `POST  ` `projects.locations.dataStores.suggestionDenyListEntries.import` — Imports all SuggestionDenyListEntry for a DataStore.
- `POST  ` `projects.locations.dataStores.suggestionDenyListEntries.purge` — Permanently deletes all SuggestionDenyListEntry for a DataStore.

## `updateCompletionConfig` (1)

- `PATCH ` `projects.locations.dataStores.updateCompletionConfig` — Updates the CompletionConfigs.

## `updateDataConnector` (1)

- `PATCH ` `projects.locations.collections.updateDataConnector` — Updates a DataConnector.

## `updateDocumentProcessingConfig` (1)

- `PATCH ` `projects.locations.dataStores.updateDocumentProcessingConfig` — Updates the DocumentProcessingConfig. DocumentProcessingConfig is a singleon resource of DataStore. It's empty when DataStore is created. The first call to this method will set up DocumentProcessingConfig.

## `userEvents` (4)

- `GET   ` `projects.locations.dataStores.userEvents.collect` — Writes a single user event from the browser. This uses a GET request to due to browser restriction of POST-ing to a third-party domain. This method is used only by the Discovery Engine API JavaScript pixel and Google Tag Manager. Users should not call this method directly.
- `POST  ` `projects.locations.dataStores.userEvents.import` — Bulk import of user events. Request processing might be synchronous. Events that already exist are skipped. Use this method for backfilling historical user events. Operation.response is of type ImportResponse. Note that it is possible for a subset of the items to be successfully inserted. Operation.metadata is of type ImportMetadata.
- `POST  ` `projects.locations.dataStores.userEvents.purge` — Deletes permanently all user events specified by the filter provided. Depending on the number of events specified by the filter, this operation could take hours or days to complete. To test a filter, use the list command first.
- `POST  ` `projects.locations.dataStores.userEvents.write` — Writes a single user event.

## `userLicenses` (1)

- `GET   ` `projects.locations.userStores.userLicenses.list` — Lists the User Licenses.

## `widgetConfigs` (2)

- `GET   ` `projects.locations.dataStores.widgetConfigs.get` — Gets a WidgetConfig.
- `PATCH ` `projects.locations.dataStores.widgetConfigs.patch` — Update a WidgetConfig.

## `write` (1)

- `POST  ` `projects.locations.userEvents.write` — Writes a single user event.

