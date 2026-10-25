import { z } from 'zod';
import { SurfaceSchema } from './context.js';

/** Content-free observations of this client. Requirement versions remain strings (1.10 != 1.1). */
export const HostCapabilityProfileSchema = z
  .object({
    surface: SurfaceSchema,
    platform: z.string().max(64),
    build: z.string().max(64),
    mode: z.enum(['read', 'write', 'compose', 'unknown']),
    requirements: z.record(z.boolean()),
  })
  .strict();
export type HostCapabilityProfile = z.infer<typeof HostCapabilityProfileSchema>;
export type SupportDecision = {
  supported: boolean;
  reason?: 'unsupported_api' | 'read_only' | 'compose_required' | 'wrong_surface';
  missing?: string[];
};

/** Identifiers survive a task; native Office proxy objects never cross this boundary. */
export const NativeTargetSchema = z
  .object({
    surface: SurfaceSchema,
    documentId: z.string().min(1).max(256),
    objectId: z.string().min(1).max(256),
    parentObjectId: z.string().min(1).max(256).optional(),
    locator: z.string().min(1).max(1024),
    revision: z.string().min(1).max(256),
    scope: z.string().min(1).max(1024),
  })
  .strict();
export type NativeTarget = z.infer<typeof NativeTargetSchema>;
export interface TextTargetCapture {
  target: NativeTarget;
  text: string;
}
