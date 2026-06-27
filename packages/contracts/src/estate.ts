import { z } from 'zod';

/**
 * The cross-repository (Plane B) handle. Where a `ContextRef` points at something in the
 * active document, an `EstateRef` points at something in the user's Microsoft 365 estate —
 * a mail item, a calendar event, a OneDrive/SharePoint file, a person — reachable via
 * Microsoft Graph **as the signed-in user**. `@ge/graph-client` resolves it into the same
 * `ResolvedContext` everything else flows through, so Plane A and Plane B converge.
 */
export const EstateSourceSchema = z.enum([
  'mail', // a single Outlook message
  'mail-thread', // a conversation
  'calendar', // a calendar event
  'drive-item', // a OneDrive/SharePoint file
  'site', // a SharePoint site/list item
  'person', // a directory user/contact
]);
export type EstateSource = z.infer<typeof EstateSourceSchema>;

export const EstateRefSchema = z.object({
  source: EstateSourceSchema,
  id: z.string(), // Graph id (message id, event id, driveItem id, user id…)
  title: z.string().optional(),
  preview: z.string().optional(),
  webUrl: z.string().optional(), // deep link back to the item
  driveId: z.string().optional(), // for drive-item
});
export type EstateRef = z.infer<typeof EstateRefSchema>;
