import { describe, it, expect } from 'vitest';
import type { CapabilityManifest } from './capability.js';
import {
  ParsedCommandSchema,
  WRITE_VERB_TO_KIND,
  extractCommandBlock,
  parseCommandLine,
  parseCommandBlock,
  grammarFor,
  isCommandParseError,
} from './command-grammar.js';

const excelManifest: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['range', 'sheet'],
  actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true }],
};

const wordManifest: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document'],
  actuations: [
    { kind: 'tracked-change', surface: 'word', title: 'Insert tracked change', reversible: true },
  ],
};

describe('command-grammar — verb map', () => {
  it('maps write verbs to actuation kinds', () => {
    expect(WRITE_VERB_TO_KIND).toEqual({ set: 'write-cells', suggest: 'tracked-change' });
  });
});

describe('command-grammar — parseCommandLine (control + reads)', () => {
  it('parses no-arg verbs', () => {
    expect(parseCommandLine('outline')).toEqual({ verb: 'outline' });
    expect(parseCommandLine('done')).toEqual({ verb: 'done' });
    expect(parseCommandLine('help')).toEqual({ verb: 'help' });
  });

  it('is case-insensitive on the verb', () => {
    expect(parseCommandLine('OUTLINE')).toEqual({ verb: 'outline' });
    expect(parseCommandLine('Done')).toEqual({ verb: 'done' });
  });

  it('parses read with an Excel selector and read with no selector (Word whole-doc)', () => {
    expect(parseCommandLine('read Sales!C2:C7')).toEqual({
      verb: 'read',
      selector: 'Sales!C2:C7',
    });
    expect(parseCommandLine('read')).toEqual({ verb: 'read', selector: '' });
  });

  it('parses search text and strips wrapping quotes', () => {
    expect(parseCommandLine('search margin')).toEqual({ verb: 'search', text: 'margin' });
    expect(parseCommandLine('search "net margin"')).toEqual({
      verb: 'search',
      text: 'net margin',
    });
    expect(parseCommandLine('search')).toMatchObject({ error: expect.stringContaining('search') });
  });
});

describe('command-grammar — set quoting (value is the full remainder)', () => {
  it('keeps a value with spaces and commas verbatim (formula)', () => {
    expect(parseCommandLine('set Sales!F2 =SUM(A1, A2)')).toEqual({
      verb: 'set',
      cell: 'Sales!F2',
      value: '=SUM(A1, A2)',
    });
  });

  it('keeps a multi-word literal value verbatim', () => {
    expect(parseCommandLine('set B16 Total Revenue by Region')).toEqual({
      verb: 'set',
      cell: 'B16',
      value: 'Total Revenue by Region',
    });
  });

  it('errors when set is missing a value', () => {
    expect(parseCommandLine('set A1')).toMatchObject({
      error: expect.stringContaining('cell and a value'),
    });
  });
});

describe('command-grammar — suggest quoting (embedded quotes, escapes, separators)', () => {
  it('parses two quoted strings separated by =>', () => {
    expect(parseCommandLine('suggest "old text" => "new text"')).toEqual({
      verb: 'suggest',
      oldText: 'old text',
      newText: 'new text',
    });
  });

  it('tolerates -> and tight/loose spacing around the separator', () => {
    expect(parseCommandLine('suggest "a"->"b"')).toEqual({
      verb: 'suggest',
      oldText: 'a',
      newText: 'b',
    });
    expect(parseCommandLine('suggest "a"    =>    "b"')).toEqual({
      verb: 'suggest',
      oldText: 'a',
      newText: 'b',
    });
  });

  it('honors escaped quotes inside the strings', () => {
    expect(
      parseCommandLine('suggest "the \\"SLA\\" is 99.5%" => "the SLA is ~99.5% (source needed)"'),
    ).toEqual({
      verb: 'suggest',
      oldText: 'the "SLA" is 99.5%',
      newText: 'the SLA is ~99.5% (source needed)',
    });
  });

  it('honors escaped backslashes', () => {
    expect(parseCommandLine('suggest "a\\\\b" => "c"')).toEqual({
      verb: 'suggest',
      oldText: 'a\\b',
      newText: 'c',
    });
  });

  it('errors on a malformed suggest (missing separator / unterminated quote)', () => {
    expect(parseCommandLine('suggest "old" "new"')).toMatchObject({
      error: expect.stringContaining('two quoted strings'),
    });
    expect(parseCommandLine('suggest "old => new')).toMatchObject({
      error: expect.stringContaining('two quoted strings'),
    });
  });
});

describe('command-grammar — corrective errors + did-you-mean', () => {
  it('suggests the nearest verb for a typo', () => {
    expect(parseCommandLine('sett A1 5')).toMatchObject({
      error: expect.stringMatching(/unknown verb "sett".*did you mean "set"/),
    });
    expect(parseCommandLine('serch foo')).toMatchObject({
      error: expect.stringContaining('did you mean "search"'),
    });
    expect(parseCommandLine('outln')).toMatchObject({
      error: expect.stringContaining('did you mean "outline"'),
    });
  });

  it('omits a suggestion for a totally unrelated token', () => {
    const res = parseCommandLine('frobnicate everything');
    expect(isCommandParseError(res)).toBe(true);
    if (isCommandParseError(res)) {
      expect(res.error).toContain('unknown verb');
      expect(res.error).not.toContain('did you mean');
      expect(res.error).toContain('run help');
    }
  });
});

describe('command-grammar — fence extraction', () => {
  it('extracts the cmd block ignoring the thought preamble and trailing prose', () => {
    const text = [
      '**thought** I should read the totals first, then write them back.',
      '',
      '```cmd',
      'read Sales!C2:C7',
      'set Sales!F2 =SUM(C2:C7)',
      '```',
      '',
      'That should do it.',
    ].join('\n');
    expect(extractCommandBlock(text)).toBe('read Sales!C2:C7\nset Sales!F2 =SUM(C2:C7)');
  });

  it('returns null when there is no cmd fence', () => {
    expect(extractCommandBlock('**thought** still thinking, no commands yet.')).toBeNull();
    // A non-cmd fence (e.g. ```answer) is not a command block.
    expect(extractCommandBlock('```answer\nall done\n```')).toBeNull();
  });
});

describe('command-grammar — parseCommandBlock (skips comments + blanks)', () => {
  it('parses each line, skipping comments and blank lines', () => {
    const text = ['```cmd', '# read the region totals', '', 'read Sales!C2:C7', 'done', '```'].join(
      '\n',
    );
    const { found, commands } = parseCommandBlock(text);
    expect(found).toBe(true);
    expect(commands).toEqual([{ verb: 'read', selector: 'Sales!C2:C7' }, { verb: 'done' }]);
  });

  it('reports found=false with no fence (re-prompt, not an error)', () => {
    const { found, commands } = parseCommandBlock('**thought** thinking…');
    expect(found).toBe(false);
    expect(commands).toEqual([]);
  });

  it('collects a corrective error inline among good commands', () => {
    const text = ['```cmd', 'read A1', 'writ A1 5', '```'].join('\n');
    const { commands } = parseCommandBlock(text);
    expect(commands[0]).toEqual({ verb: 'read', selector: 'A1' });
    expect(isCommandParseError(commands[1]!)).toBe(true);
  });
});

describe('command-grammar — capability scoping', () => {
  it('Excel advertises set (write-cells) and an A1 read selector, NOT suggest', () => {
    const verbs = grammarFor(excelManifest).map((v) => v.verb);
    expect(verbs).toContain('set');
    expect(verbs).not.toContain('suggest');
    const read = grammarFor(excelManifest).find((v) => v.verb === 'read');
    expect(read?.usage).toBe('read <A1|NamedRange>');
  });

  it('Word advertises suggest (tracked-change) and a whole-doc read, NOT set', () => {
    const verbs = grammarFor(wordManifest).map((v) => v.verb);
    expect(verbs).toContain('suggest');
    expect(verbs).not.toContain('set');
    const read = grammarFor(wordManifest).find((v) => v.verb === 'read');
    expect(read?.usage).toBe('read');
  });

  it('always advertises read/control verbs regardless of actuations', () => {
    const noWrite: CapabilityManifest = {
      surface: 'powerpoint',
      contextKinds: ['slide'],
      actuations: [],
    };
    const verbs = grammarFor(noWrite).map((v) => v.verb);
    expect(verbs).toEqual(expect.arrayContaining(['outline', 'read', 'search', 'done', 'help']));
    expect(verbs).not.toContain('set');
    expect(verbs).not.toContain('suggest');
  });
});

describe('command-grammar — ParsedCommandSchema validates parser output', () => {
  it('round-trips every command shape', () => {
    for (const cmd of [
      parseCommandLine('outline'),
      parseCommandLine('read A1'),
      parseCommandLine('search foo'),
      parseCommandLine('set A1 =1+1'),
      parseCommandLine('suggest "a" => "b"'),
      parseCommandLine('done'),
      parseCommandLine('help'),
    ]) {
      expect(isCommandParseError(cmd)).toBe(false);
      expect(() => ParsedCommandSchema.parse(cmd)).not.toThrow();
    }
  });
});
