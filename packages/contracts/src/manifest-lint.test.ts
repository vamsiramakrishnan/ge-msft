import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { lintUnifiedManifest, lintOneNoteManifest } from './manifest-lint.js';

/**
 * Manifest conformance: lint the REAL shipped manifests (must be error-free) and exercise the
 * negative cases against synthetic broken manifests, so the lint is proven to actually fail on the
 * regressions it claims to catch. Files are read from the repo root (vitest cwd).
 */

const repoRoot = process.cwd();
const unified = () =>
  readFileSync(resolve(repoRoot, 'manifests/m365-unified.manifest.json'), 'utf8');
const onenote = () => readFileSync(resolve(repoRoot, 'manifests/onenote.manifest.xml'), 'utf8');

const errors = (fs: { level: string }[]): { level: string }[] =>
  fs.filter((f) => f.level === 'error');

describe('unified manifest (real)', () => {
  it('lints clean with zero errors', () => {
    const findings = lintUnifiedManifest(unified());
    expect(errors(findings)).toEqual([]);
  });

  it('has no warnings either (all surface scopes present, placeholders clean)', () => {
    expect(lintUnifiedManifest(unified())).toEqual([]);
  });
});

describe('unified manifest (negative cases)', () => {
  it('flags invalid JSON', () => {
    const fs = lintUnifiedManifest('{ not json');
    expect(fs.some((f) => f.code === 'invalid-json')).toBe(true);
  });

  it('flags a ribbon control wired to an unknown actionId', () => {
    const m = JSON.parse(unified());
    m.extensions[0].ribbons[0].tabs[0].groups[0].controls[0].actionId = 'doesNotExist';
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'unwired-action')).toBe(true);
  });

  it('flags a ribbon control missing the Mac-required 80px icon', () => {
    const m = JSON.parse(unified());
    m.extensions[0].ribbons[0].tabs[0].groups[0].controls[0].icons = [
      { size: 16, url: 'icon-16.png' },
      { size: 32, url: 'icon-32.png' },
    ];
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'missing-ribbon-icon')).toBe(true);
  });

  it('flags a ribbon group missing the Mac-required icons', () => {
    const m = JSON.parse(unified());
    m.extensions[0].ribbons[0].tabs[0].groups[0].icons = [];
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'missing-ribbon-group-icon')).toBe(true);
  });

  it('flags an on-send event that is not a softBlock gate', () => {
    const m = JSON.parse(unified());
    m.extensions[0].autoRunEvents[0].events[0].options.sendMode = 'promptUser';
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'onsend-not-softblock')).toBe(true);
  });

  it('flags a half-edited (lowercase) REPLACE_ placeholder', () => {
    const m = JSON.parse(unified());
    m.validDomains = ['REPLACE_web_shell_domain'];
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'malformed-placeholder')).toBe(true);
  });

  it('flags a missing surface scope', () => {
    const m = JSON.parse(unified());
    m.extensions[0].requirements.scopes = ['mail', 'workbook', 'document']; // drop presentation
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'missing-scope')).toBe(true);
  });
});

describe('onenote manifest (real)', () => {
  it('lints clean with zero errors', () => {
    expect(errors(lintOneNoteManifest(onenote()))).toEqual([]);
  });
});

describe('onenote manifest (negative cases)', () => {
  it('flags a broadened permission', () => {
    const xml = onenote().replace('ReadWriteDocument', 'ReadWriteMailbox');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'bad-permission')).toBe(true);
  });

  it('flags an insecure (http) AppDomain', () => {
    const xml = onenote().replace(
      '<AppDomain>https://REPLACE_WEB_SHELL_DOMAIN</AppDomain>',
      '<AppDomain>http://REPLACE_WEB_SHELL_DOMAIN</AppDomain>',
    );
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'insecure-domain')).toBe(true);
  });

  it('flags a missing Notebook host', () => {
    const xml = onenote().replace('<Host Name="Notebook" />', '');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'missing-host')).toBe(true);
  });
});

describe('client-direct: no gateway token in any manifest (ADR-0001)', () => {
  // There is no gateway tier; the only origin the add-in trusts is the web-shell. Guard against a
  // REPLACE_GATEWAY_DOMAIN (or any GATEWAY-bearing token) creeping back into either package.
  it('the unified manifest contains no GATEWAY token', () => {
    expect(unified()).not.toMatch(/GATEWAY/);
  });

  it('the onenote manifest contains no GATEWAY token', () => {
    expect(onenote()).not.toMatch(/GATEWAY/);
  });
});
