import { describe, expect, it } from 'vitest';
import { extractDirectCommandProgram } from './direct-command.js';

describe('extractDirectCommandProgram', () => {
  it('keeps document reading prose on the natural-language route', () => {
    for (const text of [
      'read the entire doc and summarise it',
      'Read all 200 slides and summarize them',
      'search the document for risks',
      'outline the major findings',
    ]) {
      expect(extractDirectCommandProgram(text)).toBeUndefined();
    }
    expect(extractDirectCommandProgram('read document:40')).toBe('read document:40');
    expect(extractDirectCommandProgram('read Sales!A1:B4')).toBe('read Sales!A1:B4');
    expect(extractDirectCommandProgram('```cmd\nsearch termination\n```')).toBe(
      'search termination',
    );
  });
  it('extracts a pasted schedule command run and ignores prose plus trailing slash chat', () => {
    const program = extractDirectCommandProgram(`
Populate a mock schedule for this please
set 'Daily schedule'!B2 "Time"
set 'Daily schedule'!C2 "Monday"
set 'Daily schedule'!G12 "Wrap Up & Planning"
/summarize @this Summarize the selected range.
`);

    expect(program).toBe(
      [
        `set 'Daily schedule'!B2 "Time"`,
        `set 'Daily schedule'!C2 "Monday"`,
        `set 'Daily schedule'!G12 "Wrap Up & Planning"`,
      ].join('\n'),
    );
  });

  it('does not turn explanatory prose containing a command-looking word into CLI', () => {
    expect(extractDirectCommandProgram('why did set Daily schedule not write?')).toBeUndefined();
    expect(extractDirectCommandProgram('set a reminder for lunch')).toBeUndefined();
  });

  it('routes natural-language set requests to the planner, including one inside a command run', () => {
    for (const raw of [
      'Set A21 to =WEBSERVICE("http://intranet/api")',
      'Set A21 to =SUM(B2:B11)',
      'set A21 Total',
      'read Sheet1!A1:B11\nSet A21 to =SUM(B2:B11)',
      'read Sales!A1:B11\nset A20 5\nSet A21 to =SUM(B2:B11)',
      'Set A21 to =SUM(B2:B11)\nread Sales!A1:B11\nset A20 5',
    ]) {
      expect(extractDirectCommandProgram(raw), raw).toBeUndefined();
    }
    for (const raw of [
      'set A21 "to"',
      'set A21 true',
      'set A21 42',
      'set A21 =SUM(B2, B11)',
      'set A21 $total',
    ]) {
      expect(extractDirectCommandProgram(raw), raw).toBe(raw);
    }
  });

  it('treats a sentence that starts with a CLI verb as prose, not a command', () => {
    for (const raw of [
      'Read the last row for me',
      'read the last row for me',
      'Search for the payment clause',
      'List the open risks please',
      'Outline the key points of this deck',
    ]) {
      expect(extractDirectCommandProgram(raw), raw).toBeUndefined();
    }
    for (const cli of [
      'read Sales!A1:B9',
      "read 'Daily schedule'!B2:I10",
      'read selection',
      'list shape',
      'chart bar A1:B5',
      'inspect result:abc path=/rows offset=0 limit=20',
    ]) {
      expect(extractDirectCommandProgram(cli), cli).toBe(cli);
    }
  });

  it('treats "find and replace …" as a request, not the workspace find command', () => {
    expect(
      extractDirectCommandProgram('Find and replace Contoso-AI with ContosoAI'),
    ).toBeUndefined();
    expect(extractDirectCommandProgram('find /workspace *.md')).toBe('find /workspace *.md');
  });

  it('accepts an explicit cmd fence even when it contains one command', () => {
    expect(
      extractDirectCommandProgram(`
\`\`\`cmd
set 'Daily schedule'!B2 "Time"
\`\`\`
`),
    ).toBe(`set 'Daily schedule'!B2 "Time"`);
  });
});
