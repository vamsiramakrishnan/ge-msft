import { z } from 'zod';

/**
 * The verbs the assistant handles. `assist` is grounded chat over the unit
 * (streamAssist); the rest map to specialist agents the engine routes to.
 */
export const IntentSchema = z.enum([
  'assist', // grounded chat over the unit (StreamAssist)
  'review', // inline review pass -> Finding[] (A2A Review agent)
  'resolve-comment', // edit + reply + resolve a comment
  'regen-clause', // rewrite one content control
  'draft-slides', // generate slides from the unit
  'synthesize', // OneNote page synthesis from the notebook
  'meeting-notes', // Teams live notes + action items
]);

export type Intent = z.infer<typeof IntentSchema>;
