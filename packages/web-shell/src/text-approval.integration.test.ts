import { describe, expect, it, vi } from 'vitest';
import { type ActuationRequest } from '@ge/contracts';
import type { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, type DocBridge } from '@ge/runtime';
import { PanelController } from './controller.js';
import { renderCommandLine, textActuationTargetLabel } from './render-command.js';

describe('actual native text targets survive the controller approval boundary', () => {
  it.each([
    {
      kind: 'insert-text' as const,
      surface: 'word' as const,
      program: '/insert-text position=end text="Body append"',
      label: 'End of document body',
      params: { position: 'end', text: 'Body append' },
    },
    {
      kind: 'insert-text' as const,
      surface: 'word' as const,
      program:
        '/insert-text position=start match="Reporting clause" contextHint="Section 7" text="Anchor-relative insert"',
      label: 'Start of matched text "Reporting clause"',
      params: {
        position: 'start',
        target: { matchText: 'Reporting clause', contextHint: 'Section 7' },
        text: 'Anchor-relative insert',
      },
    },
    {
      kind: 'fill-content-control' as const,
      surface: 'word' as const,
      program: '/fill-content-control contentControlId=17 text="Memo"',
      label: 'Word content control "17"',
      params: { target: { contentControlId: '17' }, text: 'Memo' },
    },
    {
      kind: 'set-shape-text' as const,
      surface: 'powerpoint' as const,
      program: '/set-shape-text slideId=slide-1 shapeId=shape-2 text="Metrics"',
      label: 'PowerPoint slide "slide-1" / shape "shape-2"',
      params: { target: { slideId: 'slide-1', shapeId: 'shape-2' }, text: 'Metrics' },
    },
  ])(
    'shows the $kind actual destination before approval and dispatches that same request only after approval',
    async (input) => {
      const actuate = vi.fn(async (request: ActuationRequest) => ({
        ok: true,
        kind: request.kind,
        changeId: request.changeId,
      }));
      const bridge: DocBridge = {
        surface: input.surface,
        documentKey: () => 'document-1',
        listContext: async () => [],
        resolveContext: async () => [],
        actuate,
        getCapabilities: () => ({
          surface: input.surface,
          contextKinds: [],
          actuations: [
            { surface: input.surface, kind: input.kind, title: 'Native text', reversible: false },
          ],
        }),
      };
      const session = new AssistSession(bridge, {} as StreamAssistClient, {
        unit: { connectors: [], surfaceContext: { kind: input.surface } },
        autoAttach: [],
      });
      const controller = new PanelController(session, bridge);
      const pending = controller.runDirectCommands(input.program);
      for (let attempt = 0; attempt < 100 && !controller.getState().pendingPlan; attempt++)
        await new Promise((resolve) => setTimeout(resolve, 1));
      const effect = controller.getState().pendingPlan?.effects[0];
      expect(effect?.request.params).toMatchObject(input.params);
      expect(textActuationTargetLabel(effect!.request)).toContain(input.label);
      expect(renderCommandLine(effect!.request)).toContain(`/${input.kind}`);
      expect(actuate).not.toHaveBeenCalled();
      controller.approvePlan();
      await pending;
      expect(actuate).toHaveBeenCalledOnce();
      expect(actuate.mock.calls[0]![0].params).toMatchObject(input.params);
      session.dispose();
    },
  );
});
