import { z } from 'zod';

/**
 * The verbs the gateway routes. `assist` is grounded chat over the unit
 * (StreamAssist); the rest dispatch to specialist A2A agents.
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
