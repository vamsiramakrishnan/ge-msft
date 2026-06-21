import { z } from 'zod';

/** A federated/ingested enterprise data source attached to the unit. */
export const ConnectorRefSchema = z.object({
  type: z.enum(['sharepoint', 'onedrive']),
  mode: z.enum(['federated', 'ingestion']), // prefer 'federated' for ad-hoc sources
  scope: z.string().optional(), // e.g. "sites/DealRoom"; omit for all the user can see
});
export type ConnectorRef = z.infer<typeof ConnectorRefSchema>;

/** What the user is actively working on, per host surface. */
export const SurfaceContextSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('word'),
    selection: z.string().optional(),
    bodyOoxml: z.string().optional(),
  }),
  z.object({
    kind: z.literal('excel'),
    range: z.string().optional(),
    values: z.array(z.array(z.string())).optional(),
  }),
  z.object({ kind: z.literal('powerpoint'), slideText: z.string().optional() }),
  z.object({
    kind: z.literal('onenote'),
    pageId: z.string().optional(),
    sources: z.array(z.string()).optional(),
  }),
  z.object({ kind: z.literal('teams'), transcriptWindow: z.string().optional() }),
  z.object({
    kind: z.literal('outlook'),
    subject: z.string().optional(),
    body: z.string().optional(),
    from: z.string().optional(),
  }),
]);
export type SurfaceContext = z.infer<typeof SurfaceContextSchema>;

/** The composable research unit the agent grounds on. */
export const UnitDescriptorSchema = z.object({
  notebookId: z.string().optional(), // the curated NotebookLM core (precision)
  connectors: z.array(ConnectorRefSchema), // the live federated edge (breadth)
  restrictToNotebook: z.boolean().optional(), // true => answer only from the notebook
  surfaceContext: SurfaceContextSchema,
});
export type UnitDescriptor = z.infer<typeof UnitDescriptorSchema>;
