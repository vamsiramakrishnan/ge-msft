import { z } from 'zod';

/**
 * Lenient Zod views of the Microsoft Graph resources we read. Partial + passthrough: Graph
 * returns far more; we depend only on these fields and never break on extras.
 */

export const GraphEmailAddressSchema = z
  .object({ address: z.string().optional(), name: z.string().optional() })
  .passthrough();

export const GraphRecipientSchema = z
  .object({ emailAddress: GraphEmailAddressSchema.optional() })
  .passthrough();

export const GraphItemBodySchema = z
  .object({ contentType: z.string().optional(), content: z.string().optional() })
  .passthrough();

export const GraphMessageSchema = z
  .object({
    id: z.string().optional(),
    subject: z.string().optional(),
    from: GraphRecipientSchema.optional(),
    toRecipients: z.array(GraphRecipientSchema).optional(),
    receivedDateTime: z.string().optional(),
    body: GraphItemBodySchema.optional(),
    bodyPreview: z.string().optional(),
    webLink: z.string().optional(),
  })
  .passthrough();
export type GraphMessage = z.infer<typeof GraphMessageSchema>;

export const GraphEventSchema = z
  .object({
    id: z.string().optional(),
    subject: z.string().optional(),
    bodyPreview: z.string().optional(),
    body: GraphItemBodySchema.optional(),
    start: z.object({ dateTime: z.string().optional() }).passthrough().optional(),
    end: z.object({ dateTime: z.string().optional() }).passthrough().optional(),
    location: z.object({ displayName: z.string().optional() }).passthrough().optional(),
    organizer: GraphRecipientSchema.optional(),
    attendees: z.array(GraphRecipientSchema).optional(),
    webLink: z.string().optional(),
  })
  .passthrough();
export type GraphEvent = z.infer<typeof GraphEventSchema>;

export const GraphDriveItemSchema = z
  .object({
    id: z.string().optional(),
    name: z.string().optional(),
    webUrl: z.string().optional(),
    size: z.number().optional(),
    file: z.object({ mimeType: z.string().optional() }).passthrough().optional(),
    parentReference: z.object({ driveId: z.string().optional() }).passthrough().optional(),
  })
  .passthrough();
export type GraphDriveItem = z.infer<typeof GraphDriveItemSchema>;

/** One `GET .../children` page over the `/shared` app folder — just enough to list files. */
export const GraphChildrenSchema = z
  .object({
    value: z
      .array(
        z
          .object({
            name: z.string(),
            size: z.number().optional(),
            file: z.object({}).passthrough().optional(),
          })
          .passthrough(),
      )
      .default([]),
  })
  .passthrough();
export type GraphChildren = z.infer<typeof GraphChildrenSchema>;

export const GraphUserSchema = z
  .object({
    id: z.string().optional(),
    displayName: z.string().optional(),
    mail: z.string().optional(),
    userPrincipalName: z.string().optional(),
    jobTitle: z.string().optional(),
  })
  .passthrough();
export type GraphUser = z.infer<typeof GraphUserSchema>;

/** /search/query response: hitsContainers[].hits[].resource is the underlying entity. */
export const GraphSearchResponseSchema = z
  .object({
    value: z
      .array(
        z
          .object({
            hitsContainers: z
              .array(
                z
                  .object({
                    hits: z
                      .array(
                        z
                          .object({
                            hitId: z.string().optional(),
                            summary: z.string().optional(),
                            resource: z.record(z.unknown()).optional(),
                          })
                          .passthrough(),
                      )
                      .optional(),
                  })
                  .passthrough(),
              )
              .optional(),
          })
          .passthrough(),
      )
      .optional(),
  })
  .passthrough();
export type GraphSearchResponse = z.infer<typeof GraphSearchResponseSchema>;
