import { asChangeId, type Surface, type ActuationKind } from '@ge/contracts';
import type { DurableProgramRecord } from '@ge/runtime';
import type { SavedOfficeTasksState } from '../controller.js';

/** Explicit view fixtures only. These records are never inserted into a runtime task store. */
export function savedOfficeTaskPreview(
  params: URLSearchParams,
  surface: Surface,
): SavedOfficeTasksState | undefined {
  const mode = params.get('saved-tasks');
  if (!mode) return undefined;
  if (mode === 'empty') return { available: true, loaded: true, loading: false, items: [] };
  if (mode === 'loading') return { available: true, loaded: false, loading: true, items: [] };
  if (mode === 'error')
    return {
      available: true,
      loaded: false,
      loading: false,
      items: [],
      error: 'The authenticated account changed. Reopen the add-in before continuing.',
    };
  const blocked = mode === 'blocked';
  const verified = mode === 'complete';
  const kind: ActuationKind =
    surface === 'powerpoint'
      ? 'set-shape-text'
      : surface === 'excel'
        ? 'write-cells'
        : 'fill-content-control';
  const program =
    surface === 'powerpoint'
      ? '/set-shape-text slideId="slide-1" shapeId="shape-1" text="Reviewed close"\nfinish when=verified'
      : surface === 'excel'
        ? 'set Summary!A1 "Reviewed close"\nfinish when=verified'
        : '/fill-content-control contentControlId="7" text="Reviewed close"\nfinish when=verified';
  const record: DurableProgramRecord = {
    version: 1 as const,
    program: {
      version: 1 as const,
      id: 'month-close-memo',
      surface,
      stages: [{ id: 'memo', dependsOn: [], program }],
      evidence: { kind: 'literal' as const },
    },
    fingerprint: 'sha256:' + 'a'.repeat(64),
    updatedAt: '2026-10-25T00:00:00Z',
    stages: [
      {
        id: 'memo',
        state: verified
          ? ('verified' as const)
          : blocked
            ? ('blocked' as const)
            : ('pending' as const),
        inspections: [],
        ...(verified
          ? {
              effect: { changeId: 'preview-effect', kind, fingerprint: 'sha256:' + 'b'.repeat(64) },
              receipt: {
                ok: true,
                kind,
                changeId: asChangeId('preview-effect'),
                verification: {
                  status: 'verified' as const,
                  afterHash: 'sha256:' + 'c'.repeat(64),
                },
              },
            }
          : {}),
      },
    ],
  };
  return {
    available: true,
    loaded: true,
    loading: false,
    items: [
      {
        id: record.program.id,
        surface,
        updatedAt: record.updatedAt,
        stageCount: 1,
        verifiedStages: verified ? 1 : 0,
        state: verified ? 'verified' : blocked ? 'blocked' : 'ready',
        durable: true,
      },
    ],
    selected: {
      record,
      canResume: !blocked && !verified,
      blockers: blocked
        ? [
            'A previous native change has an unknown outcome. Inspect it in Office before continuing.',
          ]
        : [],
      stages: [
        {
          id: 'memo',
          status: blocked ? 'blocked' : verified ? 'current' : 'pending',
          ...(blocked ? { reason: 'Unknown outcomes are never replayed automatically.' } : {}),
        },
      ],
    },
  };
}
