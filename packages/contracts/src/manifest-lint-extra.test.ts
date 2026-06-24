import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { lintUnifiedManifest, lintOneNoteManifest } from './manifest-lint.js';

/**
 * Additional negative-case coverage for the offline manifest lint, hitting the branches the
 * existing suite does not: missing top-level keys, the missing-extension early return,
 * bad capability entries, no-runtime-actions, the autoRunEvent unwired-action path, and the
 * OneNote well-formedness / id / source / empty failures.
 */

const repoRoot = process.cwd();
const unified = (): string =>
  readFileSync(resolve(repoRoot, 'manifests/m365-unified.manifest.json'), 'utf8');
const onenote = (): string =>
  readFileSync(resolve(repoRoot, 'manifests/onenote.manifest.xml'), 'utf8');

const codes = (fs: { code: string }[]): string[] => fs.map((f) => f.code);

describe('unified manifest — structural failures', () => {
  it('flags every missing top-level key', () => {
    const fs = lintUnifiedManifest('{}');
    const cs = codes(fs);
    for (const key of ['$schema', 'manifestVersion', 'id', 'name', 'developer', 'validDomains']) {
      expect(fs.some((f) => f.code === 'missing-key' && f.message.includes(key))).toBe(true);
    }
    // Empty object also has no extensions block.
    expect(cs).toContain('missing-extension');
  });

  it('flags a missing name.short even when name is present as an object', () => {
    const m = JSON.parse(unified());
    delete m.name.short;
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'missing-key' && f.message.includes('name.short'))).toBe(true);
  });

  it('returns early with missing-extension when extensions is empty, skipping wiring checks', () => {
    const m = JSON.parse(unified());
    m.extensions = [];
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'missing-extension')).toBe(true);
    // Because of the early return, no action-wiring findings are produced.
    expect(fs.some((f) => f.code === 'unwired-action')).toBe(false);
    expect(fs.some((f) => f.code === 'no-runtime-actions')).toBe(false);
  });

  it('flags a capability entry lacking name/minVersion', () => {
    const m = JSON.parse(unified());
    m.extensions[0].requirements.capabilities = [{ name: 'Mailbox' }]; // no minVersion
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'bad-capability')).toBe(true);
  });

  it('flags when no runtime actions are declared at all', () => {
    const m = JSON.parse(unified());
    for (const rt of m.extensions[0].runtimes ?? []) rt.actions = [];
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'no-runtime-actions')).toBe(true);
  });

  it('flags an autoRunEvent wired to an unknown actionId', () => {
    const m = JSON.parse(unified());
    m.extensions[0].autoRunEvents[0].events[0].actionId = 'ghostAction';
    const fs = lintUnifiedManifest(JSON.stringify(m));
    expect(fs.some((f) => f.code === 'unwired-action' && f.message.includes('autoRunEvent'))).toBe(
      true,
    );
  });
});

describe('onenote manifest — well-formedness failures', () => {
  it('flags an empty file', () => {
    const fs = lintOneNoteManifest('   ');
    expect(fs).toEqual([{ level: 'error', code: 'empty', message: expect.any(String) }]);
  });

  it('flags a missing/duplicated OfficeApp root', () => {
    const fs = lintOneNoteManifest('<NotApp></NotApp>');
    expect(fs.some((f) => f.code === 'bad-root')).toBe(true);
  });

  it('flags a missing <Id> element', () => {
    const xml = onenote().replace(/<Id>[^<]+<\/Id>/, '');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'missing-id')).toBe(true);
  });

  it('flags a missing <SourceLocation>', () => {
    const xml = onenote().replace(/<SourceLocation\b/g, '<NotSourceLocation ');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'missing-source')).toBe(true);
  });

  it('flags a non-TaskPaneApp OfficeApp', () => {
    const xml = onenote().replace('xsi:type="TaskPaneApp"', 'xsi:type="ContentApp"');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'not-taskpane')).toBe(true);
  });

  it('flags a malformed placeholder inside the XML (raw scan path)', () => {
    const xml = onenote().replace(/REPLACE_[A-Z0-9_]+/, 'REPLACE_lowercase_token');
    const fs = lintOneNoteManifest(xml);
    expect(fs.some((f) => f.code === 'malformed-placeholder')).toBe(true);
  });
});
