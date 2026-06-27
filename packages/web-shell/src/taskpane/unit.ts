import type { Surface, UnitDescriptor, SurfaceContext } from '@ge/contracts';

/** An empty `SurfaceContext` for the surface — the live host content is attached lazily as refs. */
export function emptySurfaceContext(surface: Surface): SurfaceContext {
  switch (surface) {
    case 'word':
      return { kind: 'word' };
    case 'excel':
      return { kind: 'excel' };
    case 'powerpoint':
      return { kind: 'powerpoint' };
    case 'onenote':
      return { kind: 'onenote' };
    case 'teams':
      return { kind: 'teams' };
    case 'outlook':
      return { kind: 'outlook' };
    default: {
      // Exhaustiveness guard: a new Surface must be handled above.
      const never: never = surface;
      throw new Error(`Unhandled surface: ${String(never)}`);
    }
  }
}

export interface InitialUnitOptions {
  surface: Surface;
  notebookId?: string;
}

/**
 * Seed the research unit the agent grounds on. We start with no federated connectors (the user
 * adds SharePoint/OneDrive sources from the context tray) and an empty surface context (the
 * working document is attached as a live ref). An optional NotebookLM `notebookId` provides the
 * curated precision core.
 */
export function initialUnit(opts: InitialUnitOptions): UnitDescriptor {
  return {
    connectors: [],
    surfaceContext: emptySurfaceContext(opts.surface),
    ...(opts.notebookId ? { notebookId: opts.notebookId } : {}),
  };
}
