import { describe, expect, it } from 'vitest';
import { extractDirectCommandProgram } from './direct-command.js';

describe('extractDirectCommandProgram', () => {
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
