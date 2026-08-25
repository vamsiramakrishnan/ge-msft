import { describe, expect, it } from 'vitest';
import { multiHostOfficeXmlManifest, taskPaneXmlManifest } from './common.mjs';

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

describe('taskPaneXmlManifest — desktop ribbon bypass', () => {
  it('adds a direct task-pane ribbon command to every desktop document host', () => {
    for (const [surface, host] of [
      ['word', 'Document'],
      ['excel', 'Workbook'],
      ['powerpoint', 'Presentation'],
    ] as const) {
      const xml = taskPaneXmlManifest(cfg, surface);
      expect(xml).toContain(`Host xsi:type="${host}"`);
      expect(xml).toContain('<ExtensionPoint xsi:type="PrimaryCommandSurface">');
      expect(xml).toContain('<Group id="geminiGroup">');
      expect(xml).toContain('<Control xsi:type="Button" id="openGeminiBtn">');
      expect(xml).toContain('<Action xsi:type="ShowTaskpane">');
      expect(xml).toContain('<bt:Image size="80" resid="Icon.80" />');
      expect(xml).toContain(
        `<bt:Url id="Taskpane.Url" DefaultValue="https://shell.example/taskpane.html?host=${surface}" />`,
      );
    }
  });
});

describe('multiHostOfficeXmlManifest — centralized deployment', () => {
  it('uses one identity and declares Word, Excel, and PowerPoint with direct ribbon commands', () => {
    const xml = multiHostOfficeXmlManifest({
      ...cfg,
      officeXmlAppId: 'a212d7c5-b96f-4d7a-9a67-2c483c698b48',
      entraClientId: '7fbd455a-6fba-452e-99de-bfe4c7a7f8b2',
      webDomain: 'shell.example',
    });

    expect(xml).toContain('<Id>a212d7c5-b96f-4d7a-9a67-2c483c698b48</Id>');
    for (const host of ['Document', 'Workbook', 'Presentation']) {
      expect(xml).toContain(`<Host Name="${host}" />`);
      expect(xml).toContain(`<Host xsi:type="${host}">`);
    }
    for (const button of ['openGeminiWordBtn', 'openGeminiExcelBtn', 'openGeminiPowerPointBtn']) {
      expect(xml).toContain(`<Control xsi:type="Button" id="${button}">`);
    }
    expect(xml).toContain('<FunctionFile resid="Functions.Url" />');
    expect(xml).toContain('<WebApplicationInfo>');
    expect(xml).toContain(
      '<Resource>api://shell.example/7fbd455a-6fba-452e-99de-bfe4c7a7f8b2</Resource>',
    );
  });
});
