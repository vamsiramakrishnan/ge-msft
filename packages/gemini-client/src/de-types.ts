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

export const DeTextGroundingMetadataSchema = z
  .object({
    references: z.array(DeReferenceSchema).optional(),
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

export const DeAssistAnswerSchema = z
  .object({
    state: z.string().optional(),
    replies: z.array(DeReplySchema).optional(),
    assistSkippedReasons: z.array(z.string()).optional(),
  })
  .passthrough();

export const DeSessionInfoSchema = z.object({ session: z.string().optional() }).passthrough();

export const DeInvokedSkillSchema = z
  .object({ name: z.string().optional(), displayName: z.string().optional() })
  .passthrough();

export const DeStreamAssistResponseSchema = z
  .object({
    answer: DeAssistAnswerSchema.optional(),
    sessionInfo: DeSessionInfoSchema.optional(),
    invokedSkills: z.array(DeInvokedSkillSchema).optional(),
    assistToken: z.string().optional(),
  })
  .passthrough();

export type DeStreamAssistResponse = z.infer<typeof DeStreamAssistResponseSchema>;
