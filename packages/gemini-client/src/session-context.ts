import type { ResolvedContext, ContextValue } from '@ge/contracts';

/**
 * The live set of host objects attached to a Gemini Enterprise session. The add-in
 * adds/removes context as the user works (a selection, a slide, an email thread, a
 * SharePoint doc), and this turns the attached set into Discovery Engine
 * `query.parts[]`. Surface-agnostic: it only sees `ResolvedContext` from a bridge.
 */
export class SessionContext {
  private readonly items = new Map<string, ResolvedContext>();

  /** Attach (or replace) a resolved context object, keyed by its ref id. */
  add(ctx: ResolvedContext): void {
    this.items.set(ctx.ref.id, ctx);
  }

  remove(id: string): void {
    this.items.delete(id);
  }

  clear(): void {
    this.items.clear();
  }

  list(): ResolvedContext[] {
    return [...this.items.values()];
  }

  get size(): number {
    return this.items.size;
  }

  /** Map the attached set to Discovery Engine query parts (order preserved). */
  toQueryParts(): QueryPart[] {
    return this.list().map((c) => contextValueToQueryPart(c.value));
  }
}

/** A Discovery Engine `QueryPart` (the subset we emit). */
export type QueryPart =
  | { text: string; mimeType?: string }
  | { documentReference: { documentName: string; displayTitle?: string } }
  | { driveDocumentReference: { driveId: string; documentName?: string; displayTitle?: string } }
  | { personReference: { displayName: string; email?: string; personId?: string } };

export function contextValueToQueryPart(value: ContextValue): QueryPart {
  switch (value.as) {
    case 'text':
      return value.mimeType ? { text: value.text, mimeType: value.mimeType } : { text: value.text };
    case 'indexed-document':
      return {
        documentReference: {
          documentName: value.documentName,
          ...(value.title ? { displayTitle: value.title } : {}),
        },
      };
    case 'drive-document':
      return {
        driveDocumentReference: {
          driveId: value.driveId,
          ...(value.documentName ? { documentName: value.documentName } : {}),
          ...(value.title ? { displayTitle: value.title } : {}),
        },
      };
    case 'person':
      return {
        personReference: {
          displayName: value.displayName,
          ...(value.email ? { email: value.email } : {}),
          ...(value.personId ? { personId: value.personId } : {}),
        },
      };
  }
}
