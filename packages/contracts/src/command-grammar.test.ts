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
  reads: ['outline', 'read', 'search'],
  actuations: [
    { kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true },
    { kind: 'add-comment', surface: 'excel', title: 'Add comment', reversible: true },
    { kind: 'format-cells', surface: 'excel', title: 'Format cells', reversible: true },
  ],
};

const wordManifest: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document'],
  reads: ['outline', 'read', 'search'],
  actuations: [
    { kind: 'tracked-change', surface: 'word', title: 'Insert tracked change', reversible: true },
    { kind: 'add-comment', surface: 'word', title: 'Add comment', reversible: true },
    { kind: 'comment-reply', surface: 'word', title: 'Reply to comment', reversible: true },
  ],
};

describe('command-grammar — verb map', () => {
  it('maps write verbs to actuation kinds', () => {
    expect(WRITE_VERB_TO_KIND).toEqual({
      set: 'write-cells',
      suggest: 'tracked-change',
      comment: 'add-comment',
      format: 'format-cells',
      reply: 'comment-reply',
      slide: 'insert-slide',
      page: 'append-page',
      mail: 'reply-mail',
      post: 'post-message',
      compose: 'create-mail',
      table: 'create-table',
      chart: 'insert-chart',
      cf: 'format-conditional',
      spill: 'write-cells',
    });
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

describe('command-grammar — ADR-0005 Phase 2 effect-arg expressions', () => {
  it('a bare $var value is an expression (not a literal)', () => {
    expect(parseCommandLine('set B3 = $total')).toMatchObject({
      verb: 'set',
      cell: 'B3',
      valueExpr: { kind: 'pipeline', source: { src: 'var', name: 'total' }, stages: [] },
    });
  });

  it('a parenthesized pipeline value is an expression', () => {
    expect(parseCommandLine('set Summary!B2 = ($anz | sum Revenue)')).toMatchObject({
      verb: 'set',
      cell: 'Summary!B2',
      valueExpr: {
        kind: 'pipeline',
        source: { src: 'var', name: 'anz' },
        stages: [{ name: 'sum', args: 'Revenue' }],
      },
    });
  });

  it('a read-sourced parenthesized pipeline value is an expression', () => {
    expect(parseCommandLine('set B5 = (read Sales!A1:B9 | sum amount)')).toMatchObject({
      verb: 'set',
      valueExpr: {
        kind: 'pipeline',
        source: { src: 'read', selector: 'Sales!A1:B9' },
        stages: [{ name: 'sum', args: 'amount' }],
      },
    });
  });

  it('a literal value (formula / plain text) is NOT an expression (back-compat)', () => {
    expect(parseCommandLine('set Sales!F2 =SUM(A1, A2)')).toEqual({
      verb: 'set',
      cell: 'Sales!F2',
      value: '=SUM(A1, A2)',
    });
    expect(parseCommandLine('set B16 Total Revenue')).toEqual({
      verb: 'set',
      cell: 'B16',
      value: 'Total Revenue',
    });
  });

  it('an effect-arg expr cannot smuggle a write — the effect stays a transform STAGE (runtime rejects it)', () => {
    // `( $x | set ... )` parses structurally into a pipeline whose source is `$x` and whose
    // stage is a (non-existent) transform named `set` — it is NEVER a command. The runtime
    // evaluator rejects the `set` stage with the pure-only corrective at eval time, so it can
    // read+compute but never write. The KEY safety property here: the parser produced an
    // EXPRESSION, not a nested `set` command.
    const cmd = parseCommandLine('set B3 = ($x | set A1 1)');
    expect(cmd).toMatchObject({
      verb: 'set',
      cell: 'B3',
      valueExpr: {
        kind: 'pipeline',
        source: { src: 'var', name: 'x' },
        stages: [{ name: 'set', args: 'A1 1' }],
      },
    });
  });

  it('comment/reply text may be a $var or parenthesized pipeline; a quoted body stays literal', () => {
    expect(parseCommandLine('comment Sales!A1 $note')).toMatchObject({
      verb: 'comment',
      selector: 'Sales!A1',
      text: '',
      textExpr: { kind: 'pipeline', source: { src: 'var', name: 'note' } },
    });
    expect(parseCommandLine('reply {3f2a} ($t | count)')).toMatchObject({
      verb: 'reply',
      commentId: '{3f2a}',
      textExpr: { stages: [{ name: 'count', args: '' }] },
    });
    // A quoted body is still a literal (no textExpr).
    expect(parseCommandLine('comment Sales!A1 "anomalous spike"')).toEqual({
      verb: 'comment',
      selector: 'Sales!A1',
      text: 'anomalous spike',
    });
  });

  it('mail/post bodies may be a $var or parenthesized pipeline (composition parity)', () => {
    expect(parseCommandLine('mail ($draft | head 1)')).toMatchObject({
      verb: 'mail',
      body: '',
      bodyExpr: { stages: [{ name: 'head', args: '1' }] },
    });
    expect(parseCommandLine('post $summary')).toMatchObject({
      verb: 'post',
      text: '',
      textExpr: { kind: 'pipeline', source: { src: 'var', name: 'summary' } },
    });
    // A quoted body stays literal (no *Expr).
    expect(parseCommandLine('mail "Thanks — confirming."')).toEqual({
      verb: 'mail',
      body: 'Thanks — confirming.',
    });
  });

  it('page/compose bodies may be an expression while the title/subject stays literal', () => {
    expect(parseCommandLine('page "Notes" ($rows | count)')).toMatchObject({
      verb: 'page',
      title: 'Notes',
      body: '',
      bodyExpr: { stages: [{ name: 'count', args: '' }] },
    });
    expect(parseCommandLine('compose "Re: Q3" ($draft | head 1)')).toMatchObject({
      verb: 'compose',
      subject: 'Re: Q3',
      body: '',
      bodyExpr: { stages: [{ name: 'head', args: '1' }] },
    });
  });

  it('slide bullets may be a table expression (bulletsExpr); quoted bullets stay literal', () => {
    expect(parseCommandLine('slide "Top accounts" ($rows | select name,arr)')).toMatchObject({
      verb: 'slide',
      title: 'Top accounts',
      bullets: [],
      bulletsExpr: { stages: [{ name: 'select', args: 'name,arr' }] },
    });
    expect(parseCommandLine('slide "Q3" "Revenue up 12%" "Churn down"')).toEqual({
      verb: 'slide',
      title: 'Q3',
      bullets: ['Revenue up 12%', 'Churn down'],
    });
  });

  it('an empty ( ) expression is a corrective error', () => {
    expect(parseCommandLine('set B3 = ()')).toMatchObject({ error: expect.any(String) });
  });

  it('an unbalanced ( expression is a corrective error (not a silent literal write)', () => {
    expect(parseCommandLine('set B3 = ($a | sum amount')).toMatchObject({
      error: expect.stringContaining('unbalanced'),
    });
  });
});

describe('command-grammar — ADR-0007 host-native verbs (table / chart / cf)', () => {
  it('parses table with a bare headers flag and a name prop', () => {
    expect(parseCommandLine('table Report!A1:C12 headers name=Top')).toEqual({
      verb: 'table',
      range: 'Report!A1:C12',
      props: { headers: 'true', name: 'Top' },
    });
  });

  it('table needs a range', () => {
    expect(parseCommandLine('table')).toMatchObject({ error: expect.stringContaining('range') });
  });

  it('parses chart type + range + a quoted title (spaces preserved) + series', () => {
    expect(
      parseCommandLine('chart column Report!A1:B11 title="Top regions" series=columns'),
    ).toEqual({
      verb: 'chart',
      chartType: 'column',
      range: 'Report!A1:B11',
      props: { title: 'Top regions', series: 'columns' },
    });
  });

  it('parses chart ranges with single-quoted sheet names', () => {
    expect(parseCommandLine(`chart bar 'Project schedule'!B5:D30 title="Task Progress"`)).toEqual({
      verb: 'chart',
      chartType: 'bar',
      range: `'Project schedule'!B5:D30`,
      props: { title: 'Task Progress' },
    });
  });

  it('chart needs a type and a range', () => {
    expect(parseCommandLine('chart column')).toMatchObject({
      error: expect.stringContaining('range'),
    });
  });

  it('parses cf with an inline operator + fill', () => {
    expect(parseCommandLine('cf Sales!E2:E200 >100000 fill=#C6EFCE')).toEqual({
      verb: 'cf',
      range: 'Sales!E2:E200',
      props: { op: '>', value: '100000', fill: '#C6EFCE' },
    });
  });

  it('parses cf with a bare mode (databar) and with top=N', () => {
    expect(parseCommandLine('cf Sales!E2:E200 databar')).toEqual({
      verb: 'cf',
      range: 'Sales!E2:E200',
      props: { databar: 'true' },
    });
    expect(parseCommandLine('cf Sales!E2:E200 top=5 bottom=true')).toEqual({
      verb: 'cf',
      range: 'Sales!E2:E200',
      props: { top: '5', bottom: 'true' },
    });
  });

  it('cf needs a rule (range alone is a corrective)', () => {
    expect(parseCommandLine('cf Sales!E2:E200')).toMatchObject({
      error: expect.stringContaining('rule'),
    });
  });

  it('parses spill <range> = (<table expr>) into a valueExpr', () => {
    expect(parseCommandLine('spill Report!A1 = ($top | select Region,Revenue)')).toMatchObject({
      verb: 'spill',
      range: 'Report!A1',
      valueExpr: { kind: 'pipeline' },
    });
  });

  it('spill rejects a literal (it is the composition sink, not a verbatim writer)', () => {
    expect(parseCommandLine('spill Report!A1 = just text')).toMatchObject({
      error: expect.stringContaining('composed table'),
    });
  });

  it('spill needs a range and an expression', () => {
    expect(parseCommandLine('spill Report!A1')).toMatchObject({
      error: expect.stringContaining('range'),
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

describe('command-grammar — comment (dual form: Excel bare cell + Word quoted anchor)', () => {
  it('parses the Excel form: bare cell selector + quoted comment body', () => {
    expect(parseCommandLine('comment Sales!A16 "anomalous spike vs Q3"')).toEqual({
      verb: 'comment',
      selector: 'Sales!A16',
      text: 'anomalous spike vs Q3',
    });
  });

  it('parses the Word form: two quoted strings (anchor + comment)', () => {
    expect(parseCommandLine('comment "the SLA is 99.5%" "needs a source"')).toEqual({
      verb: 'comment',
      selector: 'the SLA is 99.5%',
      text: 'needs a source',
    });
  });

  it('honors embedded escaped quotes/backslashes in either quoted part', () => {
    expect(parseCommandLine('comment "say \\"hi\\"" "re: the \\"hi\\" claim"')).toEqual({
      verb: 'comment',
      selector: 'say "hi"',
      text: 're: the "hi" claim',
    });
    expect(parseCommandLine('comment A1 "a\\\\b path"')).toEqual({
      verb: 'comment',
      selector: 'A1',
      text: 'a\\b path',
    });
  });

  it('errors when the comment body is missing or the line is malformed', () => {
    expect(parseCommandLine('comment Sales!A16')).toMatchObject({
      error: expect.stringContaining('comment'),
    });
    expect(parseCommandLine('comment "anchor only"')).toMatchObject({
      error: expect.stringContaining('comment'),
    });
    expect(parseCommandLine('comment')).toMatchObject({
      error: expect.stringContaining('comment'),
    });
  });
});

describe('command-grammar — format (k=v pairs, values with # $ , . %)', () => {
  it('parses a range + multiple key=value pairs', () => {
    expect(
      parseCommandLine('format Sales!A16:C16 bold=true fill=#FFF2CC numberFormat=$#,##0.00'),
    ).toEqual({
      verb: 'format',
      range: 'Sales!A16:C16',
      props: { bold: 'true', fill: '#FFF2CC', numberFormat: '$#,##0.00' },
    });
  });

  it('splits each pair on the FIRST = only (values may contain =)', () => {
    expect(parseCommandLine('format A1 numberFormat=0.0%;=0')).toEqual({
      verb: 'format',
      range: 'A1',
      props: { numberFormat: '0.0%;=0' },
    });
  });

  it('errors on a format with no props or a bare key (no =)', () => {
    expect(parseCommandLine('format Sales!A16:C16')).toMatchObject({
      error: expect.stringContaining('key=value'),
    });
    expect(parseCommandLine('format')).toMatchObject({
      error: expect.stringContaining('key=value'),
    });
    expect(parseCommandLine('format A1 bold')).toMatchObject({
      error: expect.stringContaining('key=value'),
    });
  });
});

describe('command-grammar — reply (ADR-0006 comment-reply)', () => {
  it('parses a bare comment id + quoted reply body', () => {
    expect(parseCommandLine('reply {3f2a} "addressed in the redline"')).toEqual({
      verb: 'reply',
      commentId: '{3f2a}',
      text: 'addressed in the redline',
    });
  });

  it('honors embedded escaped quotes/backslashes in the reply body', () => {
    expect(parseCommandLine('reply c-12 "re: the \\"SLA\\" claim"')).toEqual({
      verb: 'reply',
      commentId: 'c-12',
      text: 're: the "SLA" claim',
    });
  });

  it('errors when the reply body is missing or the line is malformed', () => {
    expect(parseCommandLine('reply {3f2a}')).toMatchObject({
      error: expect.stringContaining('reply'),
    });
    expect(parseCommandLine('reply')).toMatchObject({
      error: expect.stringContaining('reply'),
    });
    expect(parseCommandLine('reply {3f2a} unquoted body')).toMatchObject({
      error: expect.stringContaining('reply'),
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

  it('Excel advertises comment (cell form) and format when add-comment/format-cells are present', () => {
    const specs = grammarFor(excelManifest);
    const verbs = specs.map((v) => v.verb);
    expect(verbs).toContain('comment');
    expect(verbs).toContain('format');
    expect(specs.find((v) => v.verb === 'comment')?.usage).toBe('comment <cell> "text"');
    expect(specs.find((v) => v.verb === 'format')?.usage).toBe('format <range> k=v ...');
  });

  it('Word advertises comment (quoted-anchor form) but NOT format (no format-cells)', () => {
    const specs = grammarFor(wordManifest);
    const verbs = specs.map((v) => v.verb);
    expect(verbs).toContain('comment');
    expect(verbs).not.toContain('format');
    expect(specs.find((v) => v.verb === 'comment')?.usage).toBe('comment "anchor" "text"');
  });

  it('Word advertises reply when comment-reply is in its actuations', () => {
    const specs = grammarFor(wordManifest);
    expect(specs.find((v) => v.verb === 'reply')?.usage).toBe('reply <commentId> "text"');
  });

  it('does NOT advertise reply when comment-reply is absent (Excel here has none)', () => {
    const verbs = grammarFor(excelManifest).map((v) => v.verb);
    expect(verbs).not.toContain('reply');
  });

  it('a manifest without add-comment/format-cells does NOT advertise comment/format', () => {
    const plainExcel: CapabilityManifest = {
      surface: 'excel',
      contextKinds: ['range'],
      actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write', reversible: true }],
    };
    const verbs = grammarFor(plainExcel).map((v) => v.verb);
    expect(verbs).toContain('set');
    expect(verbs).not.toContain('comment');
    expect(verbs).not.toContain('format');
  });

  it('always advertises control verbs regardless of actuations or reads', () => {
    const noWrite: CapabilityManifest = {
      surface: 'powerpoint',
      contextKinds: ['slide'],
      actuations: [],
    };
    const verbs = grammarFor(noWrite).map((v) => v.verb);
    expect(verbs).toEqual(expect.arrayContaining(['done', 'help']));
    expect(verbs).not.toContain('set');
    expect(verbs).not.toContain('suggest');
  });

  it('scopes read verbs to manifest.reads (ADR-0006): absent reads ⇒ no read verbs advertised', () => {
    const noReads: CapabilityManifest = {
      surface: 'powerpoint',
      contextKinds: ['slide'],
      actuations: [],
    };
    const verbs = grammarFor(noReads).map((v) => v.verb);
    expect(verbs).not.toContain('outline');
    expect(verbs).not.toContain('read');
    expect(verbs).not.toContain('search');
  });

  it('advertises ONLY the declared read verbs', () => {
    const someReads: CapabilityManifest = {
      surface: 'word',
      contextKinds: ['document'],
      reads: ['outline', 'search'],
      actuations: [],
    };
    const verbs = grammarFor(someReads).map((v) => v.verb);
    expect(verbs).toContain('outline');
    expect(verbs).toContain('search');
    expect(verbs).not.toContain('read'); // 'read' not declared ⇒ not advertised
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
      parseCommandLine('comment A1 "note"'),
      parseCommandLine('format A1 bold=true'),
      parseCommandLine('reply {3f2a} "ok"'),
      parseCommandLine('slide "Title" "b1" "b2"'),
      parseCommandLine('page "Notes" "body"'),
      parseCommandLine('mail "reply body"'),
      parseCommandLine('post "hello team"'),
      parseCommandLine('done'),
      parseCommandLine('help'),
    ]) {
      expect(isCommandParseError(cmd)).toBe(false);
      expect(() => ParsedCommandSchema.parse(cmd)).not.toThrow();
    }
  });
});

describe('command-grammar — ADR-0006 CLI parity verbs (slide/page/mail/post)', () => {
  it('slide: title + bullets (quote-aware)', () => {
    expect(parseCommandLine('slide "Q3" "up 12%" "churn down"')).toEqual({
      verb: 'slide',
      title: 'Q3',
      bullets: ['up 12%', 'churn down'],
    });
  });
  it('slide: title only (zero bullets)', () => {
    expect(parseCommandLine('slide "Just a title"')).toEqual({
      verb: 'slide',
      title: 'Just a title',
      bullets: [],
    });
  });
  it('slide: empty/missing title is corrective; a bare unquoted token is corrective', () => {
    expect(isCommandParseError(parseCommandLine('slide ""'))).toBe(true);
    expect(isCommandParseError(parseCommandLine('slide'))).toBe(true);
    expect(isCommandParseError(parseCommandLine('slide bareToken'))).toBe(true);
  });
  it('page: title + body', () => {
    expect(parseCommandLine('page "Meeting" "Decisions: ship"')).toEqual({
      verb: 'page',
      title: 'Meeting',
      body: 'Decisions: ship',
    });
  });
  it('page: needs exactly two quoted strings', () => {
    expect(isCommandParseError(parseCommandLine('page "only title"'))).toBe(true);
    expect(isCommandParseError(parseCommandLine('page "a" "b" "c"'))).toBe(true);
  });
  it('mail: a single quoted body', () => {
    expect(parseCommandLine('mail "Thanks — confirming."')).toEqual({
      verb: 'mail',
      body: 'Thanks — confirming.',
    });
    expect(isCommandParseError(parseCommandLine('mail'))).toBe(true);
    expect(isCommandParseError(parseCommandLine('mail ""'))).toBe(true);
  });
  it('post: a single quoted text', () => {
    expect(parseCommandLine('post "summary of decisions"')).toEqual({
      verb: 'post',
      text: 'summary of decisions',
    });
    expect(isCommandParseError(parseCommandLine('post unquoted'))).toBe(true);
  });
  it('compose: a quoted subject and body', () => {
    expect(parseCommandLine('compose "Follow-up on Q3" "Hi — summary below."')).toEqual({
      verb: 'compose',
      subject: 'Follow-up on Q3',
      body: 'Hi — summary below.',
    });
    // Subject is required; body alone (one string) or an empty subject is corrective.
    expect(isCommandParseError(parseCommandLine('compose "only subject"'))).toBe(true);
    expect(isCommandParseError(parseCommandLine('compose "" "body"'))).toBe(true);
  });
  it('escapes are honored inside quoted args', () => {
    expect(parseCommandLine('post "say \\"hi\\""')).toEqual({ verb: 'post', text: 'say "hi"' });
  });
});

describe('command-grammar — the /<kind> specialized surface (ADR-0008 §two-tier)', () => {
  it('parses /<kind> into an invoke with props + positional args', () => {
    expect(parseCommandLine('/insert-image base64=AAA alt="Q3 chart"')).toEqual({
      verb: 'invoke',
      kind: 'insert-image',
      props: { base64: 'AAA', alt: 'Q3 chart' },
      args: [],
    });
  });

  it('the command name IS the ActuationKind (case-insensitive)', () => {
    expect(parseCommandLine('/SET-PAGE-TITLE title="Q3"')).toMatchObject({
      verb: 'invoke',
      kind: 'set-page-title',
    });
  });

  it('rejects an unknown capability with a catalogue did-you-mean', () => {
    const r = parseCommandLine('/insert-imag base64=AAA');
    expect(isCommandParseError(r)).toBe(true);
    if (isCommandParseError(r)) {
      expect(r.error).toContain('unknown capability');
      expect(r.error).toContain('insert-image');
    }
  });

  it('a bare / is corrective', () => {
    expect(isCommandParseError(parseCommandLine('/'))).toBe(true);
  });
});
