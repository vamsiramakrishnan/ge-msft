import { describe, it, expect } from 'vitest';
import { toOoxmlPackage } from './actuate-plan.js';

const RELS = 'pkg:name="/_rels/.rels"';
const DOC = 'pkg:name="/word/document.xml"';

describe('toOoxmlPackage (range.insertOoxml needs a full flat-OPC package)', () => {
  it('wraps a bare paragraph in a package with a .rels part, a document and a body', () => {
    const out = toOoxmlPackage('<w:p><w:r><w:t>Summary</w:t></w:r></w:p>');
    expect(out.startsWith('<pkg:package')).toBe(true);
    expect(out).toContain(RELS);
    expect(out).toContain(DOC);
    expect(out).toContain(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Summary</w:t></w:r></w:p></w:body></w:document>',
    );
  });

  it('wraps a <w:body> or <w:document> without nesting it twice', () => {
    expect(toOoxmlPackage('<w:body><w:p/></w:body>').match(/<w:body>/g)).toHaveLength(1);
    const doc = '<w:document xmlns:w="x"><w:body><w:p/></w:body></w:document>';
    const out = toOoxmlPackage(doc);
    expect(out.match(/<w:document/g)).toHaveLength(1);
    expect(out).toContain(doc);
  });

  it('adds a missing .rels part to a package and leaves a complete one alone', () => {
    const partial =
      '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"></pkg:part></pkg:package>';
    expect(toOoxmlPackage(partial)).toContain(RELS);
    const complete = toOoxmlPackage('<w:p/>');
    expect(toOoxmlPackage(complete)).toBe(complete);
  });

  it('drops an XML declaration and returns empty for empty input', () => {
    expect(toOoxmlPackage('<?xml version="1.0"?>\n<w:p/>').startsWith('<pkg:package')).toBe(true);
    expect(toOoxmlPackage('   ')).toBe('');
  });
});
