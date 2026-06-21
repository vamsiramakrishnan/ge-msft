import { z } from 'zod';

/**
 * The capability foundation (part 1 of 2): **context capture**.
 *
 * Everything a surface can read and hand to Gemini Enterprise as live session
 * context is normalized here. A `ContextRef` is a cheap, serializable *handle* the
 * add-in surfaces in the UI ("what can I attach right now"); resolving it produces a
 * `ResolvedContext` whose `value` maps 1:1 onto a Discovery Engine `query.parts[]`
 * entry (text / indexed-document / drive-document / person). Host content is always
 * carried as **data**, never as instructions. See docs/ADR-0002-capability-model.md.
 */

export const SurfaceSchema = z.enum(['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams']);
export type Surface = z.infer<typeof SurfaceSchema>;

/** The kinds of host object the add-in can lift into a session. */
export const ContextKindSchema = z.enum([
  'selection', // the current selection in any editor
  'document', // the whole working document/body
  'paragraph', // an addressable paragraph
  'table', // a table / range of structured rows
  'range', // an Excel range
  'sheet', // an Excel worksheet
  'slide', // a PowerPoint slide
  'shape', // a shape / text box / image placeholder
  'image', // an embedded image
  'comment', // a comment / thread
  'mail-item', // an Outlook message or appointment
  'mail-thread', // an Outlook conversation
  'attachment', // a mail/file attachment
  'calendar-event', // a calendar item
  'transcript', // a Teams transcript window
  'page', // a OneNote page
  'person', // a person/contact reference
  'indexed-document', // a doc already in a connected data store (cited by name)
  'drive-document', // a Google Drive document reference
  'file', // an opaque rendered file (OOXML/PDF) for multimodal grounding
]);
export type ContextKind = z.infer<typeof ContextKindSchema>;

/** A lightweight handle to an attachable host object (what the UI lists). */
export const ContextRefSchema = z.object({
  id: z.string(), // stable within a session (e.g. "word:selection", "xl:Sheet1!A1:D9")
  kind: ContextKindSchema,
  surface: SurfaceSchema,
  title: z.string(), // human label for the chip ("Selection", "Slide 4", "Email: SLA concerns")
  preview: z.string().optional(), // short snippet for the UI
  mimeType: z.string().optional(),
  sizeBytes: z.number().optional(), // for budgeting/large-object warnings
  live: z.boolean().optional(), // true ⇒ re-resolve at send-time (e.g. "current selection")
});
export type ContextRef = z.infer<typeof ContextRefSchema>;

/**
 * The materialized value of a context object — the part the gateway-less client
 * turns into a Discovery Engine query part. Discriminated by how it grounds.
 */
export const ContextValueSchema = z.discriminatedUnion('as', [
  z.object({ as: z.literal('text'), text: z.string(), mimeType: z.string().optional() }),
  z.object({
    as: z.literal('indexed-document'),
    documentName: z.string(), // projects/.../dataStores/.../documents/...
    title: z.string().optional(),
    uri: z.string().optional(),
  }),
  z.object({
    as: z.literal('drive-document'),
    driveId: z.string(),
    documentName: z.string().optional(),
    title: z.string().optional(),
  }),
  z.object({
    as: z.literal('person'),
    displayName: z.string(),
    email: z.string().optional(),
    personId: z.string().optional(),
  }),
]);
export type ContextValue = z.infer<typeof ContextValueSchema>;

export const ResolvedContextSchema = z.object({
  ref: ContextRefSchema,
  value: ContextValueSchema,
});
export type ResolvedContext = z.infer<typeof ResolvedContextSchema>;
