import { describe, expect, it } from 'vitest';
import { capabilityModule } from '@ge/contracts';
import { WordBridge } from '@ge/bridge-word';
import { ExcelBridge } from '@ge/bridge-excel';
import { PowerPointBridge } from '@ge/bridge-powerpoint';
import { OutlookBridge } from '@ge/bridge-outlook';

describe('native module coverage follows executable handler tables', () => {
  for (const [surface, bridge] of [
    ['word', WordBridge],
    ['excel', ExcelBridge],
    ['powerpoint', PowerPointBridge],
    ['outlook', OutlookBridge],
  ] as const)
    it(`has a reviewed parameter contract for every ${surface} handler`, () => {
      for (const kind of bridge.handledActuations)
        expect(capabilityModule(surface, kind), `${surface}:${kind}`).toBeDefined();
    });
});
