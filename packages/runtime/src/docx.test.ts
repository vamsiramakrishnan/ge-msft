import { describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { textToDocx } from './docx.js';

describe('textToDocx', () => {
  it('is a valid zip holding the three Word parts, text escaped (UAT #22)', () => {
    const file = join(mkdtempSync(join(tmpdir(), 'docx-')), 'hello arjuna.docx');
    writeFileSync(file, textToDocx('test script 1\nA & B <ok>'));
    const listing = execFileSync('unzip', ['-l', file], { encoding: 'utf8' });
    expect(listing).toContain('[Content_Types].xml');
    expect(listing).toContain('_rels/.rels');
    expect(listing).toContain('word/document.xml');
    const doc = execFileSync('unzip', ['-p', file, 'word/document.xml'], { encoding: 'utf8' });
    expect(doc).toContain('<w:t xml:space="preserve">test script 1</w:t>');
    expect(doc).toContain('A &amp; B &lt;ok&gt;');
    // `unzip -t` checks every CRC.
    expect(execFileSync('unzip', ['-t', file], { encoding: 'utf8' })).toContain('No errors');
  });
});
