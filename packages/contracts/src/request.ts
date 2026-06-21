import { z } from 'zod';
import { IntentSchema } from './intent.js';
import { UnitDescriptorSchema } from './unit.js';

/**
 * Every grounded call is made as the signed-in user: the client federates the user's
 * Entra identity to Google (WIF) and sends the resulting token directly. The body:
 */
export const AssistRequestSchema = z.object({
  intent: IntentSchema,
  unit: UnitDescriptorSchema,
  query: z.string().optional(), // for 'assist'
  target: z
    .object({
      contentControlId: z.string().optional(),
      commentId: z.string().optional(),
      range: z.string().optional(),
    })
    .optional(),
  changeId: z.string().optional(), // client-generated; makes write-backs idempotent
});
export type AssistRequest = z.infer<typeof AssistRequestSchema>;

/**
 * Connector actions (upload/download/check-in-out) are never implicit in a
 * grounding request. They are distinct, explicitly-authorized calls.
 */
export const ActionRequestSchema = z.object({
  action: z.enum(['upload', 'download', 'checkout', 'checkin', 'add-page']),
  connector: z.enum(['sharepoint', 'onedrive']),
  target: z.string(), // site/path/item
  payload: z
    .object({
      filename: z.string(),
      contentBase64: z.string(),
    })
    .optional(),
  changeId: z.string(),
});
export type ActionRequest = z.infer<typeof ActionRequestSchema>;
