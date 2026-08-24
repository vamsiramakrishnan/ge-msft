import { describe, expect, it } from 'vitest';
import { taskPaneXmlManifest } from './common.mjs';

const cfg = {
  webOrigin: 'https://shell.example',
  developerName: 'Test Developer',
  supportUrl: 'https://example.com/support',
};

describe('taskPaneXmlManifest — Excel custom functions wiring', () => {
  it('the Excel manifest wires the =GE.ASK function runtime + metadata', () => {
    const xml = taskPaneXmlManifest(cfg, 'excel');
    expect(xml).toContain('<ExtendedOverrides Url="https://shell.example/functions.json" />');
    expect(xml).toContain('<FunctionFile resid="Functions.Url" />');
    expect(xml).toContain('<Host xsi:type="Workbook">');
    expect(xml).toContain(
      '<bt:Url id="Functions.Url" DefaultValue="https://shell.example/functions.html" />',
    );
  });

  it('non-Excel manifests stay free of custom-function wiring', () => {
    for (const surface of ['word', 'powerpoint'] as const) {
      const xml = taskPaneXmlManifest(cfg, surface);
      expect(xml).not.toContain('ExtendedOverrides');
      expect(xml).not.toContain('FunctionFile');
      expect(xml).toContain(`?host=${surface}`);
    }
  });
});
