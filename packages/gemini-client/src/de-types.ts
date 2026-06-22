import { z } from 'zod';

/**
 * Lenient Zod views of the Discovery Engine `v1alpha` StreamAssist wire types we
 * consume. Deliberately partial + passthrough: the engine emits many more fields
 * (see docs/api/discoveryengine/streamAssist.md) but we only depend on these, and
 * we never want an unexpected extra field to break the stream.
 */

export const DeDocumentMetadataSchema = z
  .object({
    title: z.string().optional(),
    uri: z.string().optional(),
    domain: z.string().optional(),
    pageIdentifier: z.string().optional(),
    document: z.string().optional(),
  })
  .passthrough();

export const DeReferenceSchema = z
  .object({
    content: z.string().optional(),
    documentMetadata: DeDocumentMetadataSchema.optional(),
  })
  .passthrough();

/**
 * Citation source attached to a claim. The wire shape varies (references can be
 * carried by index into `references[]`, or by inline document metadata); keep it
 * tolerant and read whatever identifies the source.
 */
export const DeCitationSourceSchema = z
  .object({
    referenceId: z.string().optional(),
    referenceIndex: z.number().optional(),
    title: z.string().optional(),
    uri: z.string().optional(),
    document: z.string().optional(),
    documentMetadata: DeDocumentMetadataSchema.optional(),
  })
  .passthrough();

/**
 * A grounded claim span. `startIndex`/`endIndex` are UTF-8 **byte** offsets into
 * the answer text (per the schema) — the stream loop converts them to UTF-16 char
 * indices before emitting. They arrive as strings on the wire (int64) but may be
 * numbers in some frames, so accept both.
 */
export const DeGroundingSupportSchema = z
  .object({
    startIndex: z.union([z.string(), z.number()]).optional(),
    endIndex: z.union([z.string(), z.number()]).optional(),
    groundingScore: z.number().optional(),
    groundingCheckRequired: z.boolean().optional(),
    sources: z.array(DeCitationSourceSchema).optional(),
  })
  .passthrough();

export const DeTextGroundingMetadataSchema = z
  .object({
    references: z.array(DeReferenceSchema).optional(),
    groundingSupports: z.array(DeGroundingSupportSchema).optional(),
  })
  .passthrough();

export const DeAssistantContentSchema = z
  .object({
    text: z.string().optional(),
    role: z.string().optional(),
    thought: z.boolean().optional(),
  })
  .passthrough();

export const DeGroundedContentSchema = z
  .object({
    content: DeAssistantContentSchema.optional(),
    textGroundingMetadata: DeTextGroundingMetadataSchema.optional(),
  })
  .passthrough();

export const DeReplySchema = z
  .object({
    replyId: z.string().optional(),
    groundedContent: DeGroundedContentSchema.optional(),
  })
  .passthrough();

/**
 * Customer policy enforcement outcome (Model Armor / banned phrases). Present only
 * when the assist call ran a policy check; a `BLOCK` verdict means the turn was
 * suppressed by policy rather than failing generically.
 */
export const DeModelArmorEnforcementResultSchema = z
  .object({
    modelArmorViolation: z.string().optional(),
    error: z
      .object({ message: z.string().optional(), code: z.number().optional() })
      .passthrough()
      .optional(),
  })
  .passthrough();

export const DePolicyEnforcementResultSchema = z
  .object({
    modelArmorEnforcementResult: DeModelArmorEnforcementResultSchema.optional(),
    bannedPhraseEnforcementResult: z
      .object({ bannedPhrases: z.array(z.string()).optional() })
      .passthrough()
      .optional(),
  })
  .passthrough();

export const DeCustomerPolicyEnforcementResultSchema = z
  .object({
    verdict: z.string().optional(),
    policyResults: z.array(DePolicyEnforcementResultSchema).optional(),
  })
  .passthrough();

export type DeGroundingSupport = z.infer<typeof DeGroundingSupportSchema>;
export type DeCitationSource = z.infer<typeof DeCitationSourceSchema>;
export type DeReference = z.infer<typeof DeReferenceSchema>;
export type DeCustomerPolicyEnforcementResult = z.infer<
  typeof DeCustomerPolicyEnforcementResultSchema
>;
export type DeReply = z.infer<typeof DeReplySchema>;

/**
 * Explicit answer/response shapes. The schemas below are annotated against these
 * so the emitted `.d.ts` stays bounded (the fully-inferred `.passthrough()` union
 * exceeds the compiler's serialization limit at the project-reference boundary).
 */
export interface DeAssistAnswer {
  state?: string;
  replies?: DeReply[];
  assistSkippedReasons?: string[];
  answerSkippedReasons?: string[];
  relatedQuestions?: string[];
  customerPolicyEnforcementResult?: DeCustomerPolicyEnforcementResult;
  [k: string]: unknown;
}

export const DeAssistAnswerSchema: z.ZodType<DeAssistAnswer> = z
  .object({
    state: z.string().optional(),
    replies: z.array(DeReplySchema).optional(),
    assistSkippedReasons: z.array(z.string()).optional(),
    answerSkippedReasons: z.array(z.string()).optional(),
    relatedQuestions: z.array(z.string()).optional(),
    customerPolicyEnforcementResult: DeCustomerPolicyEnforcementResultSchema.optional(),
  })
  .passthrough();

export const DeSessionInfoSchema = z.object({ session: z.string().optional() }).passthrough();

export const DeInvokedSkillSchema = z
  .object({ name: z.string().optional(), displayName: z.string().optional() })
  .passthrough();

export interface DeStreamAssistResponse {
  answer?: DeAssistAnswer;
  sessionInfo?: { session?: string; [k: string]: unknown };
  invokedSkills?: Array<{ name?: string; displayName?: string; [k: string]: unknown }>;
  assistToken?: string;
  [k: string]: unknown;
}

export const DeStreamAssistResponseSchema: z.ZodType<DeStreamAssistResponse> = z
  .object({
    answer: DeAssistAnswerSchema.optional(),
    sessionInfo: DeSessionInfoSchema.optional(),
    invokedSkills: z.array(DeInvokedSkillSchema).optional(),
    assistToken: z.string().optional(),
  })
  .passthrough();
