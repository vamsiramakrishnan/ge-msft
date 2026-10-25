import { describe, expect, it } from 'vitest';
import {
  BRAND,
  prodProfile,
  developmentManifest,
  multiHostOfficeXmlManifest,
  oneNoteManifest,
  outlookXmlManifest,
  releaseConfig,
  validateGeneratedManifest,
} from './common.mjs';

const env = {
  GE_BRAND: 'default',
  GE_PROD_APP_ID: '0c9a3f7e-51d2-4b8e-9f10-2a6b7c8d9e01',
  GE_PROD_OFFICE_XML_APP_ID: '1d8b4e6f-62e3-4c9f-8a21-3b7c8d9e0f12',
  GE_PROD_OUTLOOK_APP_ID: '2e7c5d8a-73f4-4dab-9b32-4c8d9e0f1a23',
  GE_PROD_ENTRA_CLIENT_ID: '3f6d6e9b-84a5-4ebc-8c43-5d9e0f1a2b34',
  GE_PROD_WEB_DOMAIN: 'addin.example.test',
  GE_PROD_DEVELOPER_NAME: 'Example Org',
  GE_PROD_WEBSITE_URL: 'https://example.test',
  GE_PROD_PRIVACY_URL: 'https://example.test/privacy',
  GE_PROD_TERMS_URL: 'https://example.test/terms',
  GE_PROD_SUPPORT_URL: 'https://example.test/support',
};

describe('production release profile', () => {
  it('builds a production-named manifest for Word, Excel, PowerPoint and Outlook by default', () => {
    const cfg = releaseConfig(prodProfile, env);
    expect(cfg.surfaces).toEqual(['word', 'excel', 'powerpoint', 'outlook']);
    // CI systems pass unset variables as empty strings.
    expect(releaseConfig(prodProfile, { ...env, GE_PROD_SURFACES: '' }).surfaces).toHaveLength(4);
    const manifest = developmentManifest(cfg);
    expect(manifest.name.short).toBe('Gemini Enterprise');
    expect(manifest.extensions[0].requirements.scopes).toEqual([
      'mail',
      'workbook',
      'document',
      'presentation',
    ]);
    expect(JSON.stringify(manifest)).not.toMatch(/Dev\b|Development/);
    expect(validateGeneratedManifest(manifest, prodProfile)).toEqual([]);
    expect(multiHostOfficeXmlManifest(cfg)).toContain(
      '<DisplayName DefaultValue="Gemini Enterprise" />',
    );
    expect(outlookXmlManifest(cfg)).toContain(
      '<DisplayName DefaultValue="Gemini Enterprise (Outlook)" />',
    );
  });

  it('drops the mail blocks when Outlook is not selected', () => {
    const cfg = releaseConfig(prodProfile, { ...env, GE_PROD_SURFACES: 'word,excel' });
    const ext = developmentManifest(cfg).extensions[0];
    expect(ext.requirements.scopes).toEqual(['workbook', 'document']);
    expect(ext.autoRunEvents).toBeUndefined();
    expect(ext.contextMenus[0].requirements.scopes).toEqual(['document', 'workbook']);
  });

  it('adds OneNote as its own XML manifest once its app id is configured', () => {
    const withOneNote = { ...env, GE_PROD_SURFACES: 'word,excel,powerpoint,outlook,onenote' };
    expect(() => releaseConfig(prodProfile, withOneNote)).toThrow(/GE_PROD_ONENOTE_APP_ID/);
    const cfg = releaseConfig(prodProfile, {
      ...withOneNote,
      GE_PROD_ONENOTE_APP_ID: '4a5e7f0c-95b6-4fcd-9d54-6e0f1a2b3c45',
    });
    expect(developmentManifest(cfg).extensions[0].requirements.scopes).not.toContain('Notebook');
    expect(oneNoteManifest(cfg)).toContain(
      '<DisplayName DefaultValue="Gemini Enterprise (OneNote)" />',
    );
  });

  it("uses the web build's Entra app when GE_PROD_ENTRA_CLIENT_ID is not set", () => {
    const { GE_PROD_ENTRA_CLIENT_ID, ...rest } = env;
    const cfg = releaseConfig(prodProfile, {
      ...rest,
      VITE_ENTRA_CLIENT_ID: GE_PROD_ENTRA_CLIENT_ID,
    });
    expect(cfg.entraClientId).toBe(GE_PROD_ENTRA_CLIENT_ID);
  });

  it('rejects unknown apps, placeholder ids, reused ids and localhost', () => {
    expect(() => releaseConfig(prodProfile, { ...env, GE_PROD_SURFACES: 'word,teams' })).toThrow(
      /unknown app/,
    );
    expect(() =>
      releaseConfig(prodProfile, {
        ...env,
        GE_PROD_APP_ID: '11111111-1111-4111-8111-111111111111',
      }),
    ).toThrow(/GE_PROD_APP_ID/);
    expect(() =>
      releaseConfig(prodProfile, {
        ...env,
        GE_PROD_OUTLOOK_APP_ID: env.GE_PROD_OFFICE_XML_APP_ID,
      }),
    ).toThrow(/must all be different/);
    expect(() => releaseConfig(prodProfile, { ...env, GE_PROD_WEB_DOMAIN: 'localhost' })).toThrow(
      /GE_PROD_WEB_DOMAIN/,
    );
    for (const domain of ['10.0.0.12', 'intranet', 'a..b.test']) {
      expect(() => releaseConfig(prodProfile, { ...env, GE_PROD_WEB_DOMAIN: domain })).toThrow(
        /GE_PROD_WEB_DOMAIN/,
      );
    }
    expect(() =>
      releaseConfig(prodProfile, { ...env, GE_PROD_PRIVACY_URL: 'http://example.test/privacy' }),
    ).toThrow(/HTTPS/);
  });

  it("keeps the build brand's development naming when no production values are set", () => {
    const manifest = developmentManifest({
      webOrigin: 'https://shell.test',
      webDomain: 'shell.test',
      appId: env.GE_PROD_APP_ID,
      entraClientId: env.GE_PROD_ENTRA_CLIENT_ID,
    });
    expect(manifest.name.short).toBe(`${BRAND.name} Dev`);
    expect(manifest.extensions[0].requirements.scopes).toHaveLength(4);
  });
});
