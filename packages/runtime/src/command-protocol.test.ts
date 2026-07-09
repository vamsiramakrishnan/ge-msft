import { describe, it, expect } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { ActuationRequestSchema } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import type { AssistRequest } from '@ge/contracts';
import { TriggerRegistry } from '@ge/triggers';
import { AssistSession, type CommandLoopEvent } from './assist-session.js';
import { compileCommand, renderCommandHelp, renderGrammarPrompt } from './command-protocol.js';
import type { DocBridge } from './bridge.js';
import { asChangeId } from '@ge/contracts';

/* ───────────────────────── compileCommand unit tests ──────────────────── */

const mint = () => asChangeId('cid-fixed');

describe('compileCommand', () => {
  it('compiles `set` → a valid write-cells ActuationRequest', () => {
    const c = compileCommand(
      { verb: 'set', cell: 'Sales!F2', value: '=C2-D2' },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'write-cells',
        surface: 'excel',
        params: { target: { range: 'Sales!F2' }, cells: [['=C2-D2']] },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `suggest` → a tracked-change with target.matchText', () => {
    const c = compileCommand(
      { verb: 'suggest', oldText: 'old', newText: 'new' },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: { kind: 'tracked-change', params: { target: { matchText: 'old' }, text: 'new' } },
    });
  });

  it('compiles `comment` → add-comment with a cell range target (Excel)', () => {
    const c = compileCommand(
      { verb: 'comment', selector: 'Sales!A16', text: 'anomalous spike' },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'add-comment',
        surface: 'excel',
        params: { target: { range: 'Sales!A16' }, text: 'anomalous spike' },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `comment` → add-comment with a matchText anchor (Word)', () => {
    const c = compileCommand(
      { verb: 'comment', selector: 'the SLA is 99.5%', text: 'needs a source' },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'add-comment',
        surface: 'word',
        params: { target: { matchText: 'the SLA is 99.5%' }, text: 'needs a source' },
      },
    });
  });

  it('compiles `format` → format-cells with typed format params', () => {
    const c = compileCommand(
      {
        verb: 'format',
        range: 'Sales!A16:C16',
        props: { bold: 'true', italic: 'false', fill: '#FFF2CC', numberFormat: '$#,##0.00' },
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'format-cells',
        surface: 'excel',
        params: {
          target: { range: 'Sales!A16:C16' },
          format: { bold: true, italic: false, fill: '#FFF2CC', numberFormat: '$#,##0.00' },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('ignores unknown format keys but errors when NO recognized prop is present', () => {
    const ok = compileCommand(
      { verb: 'format', range: 'A1', props: { bold: 'true', wibble: 'x' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(ok).toMatchObject({ kind: 'write', request: { params: { format: { bold: true } } } });

    const bad = compileCommand(
      { verb: 'format', range: 'A1', props: { wibble: 'x' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(bad).toMatchObject({ error: expect.stringContaining('recognized property') });
  });

  it('compiles `reply` → comment-reply with a commentId target (Zod-valid, changeId minted once)', () => {
    const ids: string[] = [];
    const mintOnce = () => {
      const id = asChangeId(`cid-${ids.length}`);
      ids.push(id);
      return id;
    };
    const c = compileCommand(
      { verb: 'reply', commentId: '{3f2a}', text: 'addressed in the redline' },
      { surface: 'word', mintChangeId: mintOnce },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'comment-reply',
        surface: 'word',
        params: { target: { commentId: '{3f2a}' }, text: 'addressed in the redline' },
      },
    });
    expect(ids).toHaveLength(1); // changeId minted exactly once
    if ('request' in c) {
      expect(c.request.changeId).toBe('cid-0');
      expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
    }
  });

  it('compiles reads to read intents', () => {
    expect(compileCommand({ verb: 'outline' }, { surface: 'excel', mintChangeId: mint })).toEqual({
      kind: 'read',
      intent: { read: 'outline' },
    });
    expect(
      compileCommand({ verb: 'read', selector: 'A1:B2' }, { surface: 'excel', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'range', selector: 'A1:B2' } });
    expect(
      compileCommand({ verb: 'search', text: 'x' }, { surface: 'word', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'search', text: 'x' } });
    expect(
      compileCommand({ verb: 'list', kind: 'comment' }, { surface: 'word', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'list-context', kind: 'comment' } });
    expect(
      compileCommand(
        { verb: 'inspect', selector: 'xl:Sales!A1:C9' },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'read', intent: { read: 'inspect-context', selector: 'xl:Sales!A1:C9' } });
    expect(
      compileCommand(
        { verb: 'properties', selector: 'word:selection' },
        { surface: 'word', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'read', intent: { read: 'properties', selector: 'word:selection' } });
    expect(compileCommand({ verb: 'comments' }, { surface: 'word', mintChangeId: mint })).toEqual({
      kind: 'read',
      intent: { read: 'context-kind', kind: 'comment' },
    });
    expect(
      compileCommand({ verb: 'attachments' }, { surface: 'outlook', mintChangeId: mint }),
    ).toEqual({ kind: 'read', intent: { read: 'context-kind', kind: 'attachment' } });
    expect(compileCommand({ verb: 'tables' }, { surface: 'excel', mintChangeId: mint })).toEqual({
      kind: 'read',
      intent: { read: 'context-kind', kind: 'table' },
    });
    expect(
      compileCommand(
        { verb: 'slides', selector: 's1' },
        { surface: 'powerpoint', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'read', intent: { read: 'context-kind', kind: 'slide', selector: 's1' } });
    expect(compileCommand({ verb: 'neighbors' }, { surface: 'excel', mintChangeId: mint })).toEqual(
      { kind: 'read', intent: { read: 'neighbors' } },
    );
    expect(
      compileCommand(
        { verb: 'context', hints: ['analytical', 'upload-preferred'] },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({
      kind: 'read',
      intent: { read: 'context-strategy', hints: ['analytical', 'upload-preferred'] },
    });
    expect(
      compileCommand(
        { verb: 'open', selector: 'xl:Sales!A1:C9' },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'read', intent: { read: 'open-context', selector: 'xl:Sales!A1:C9' } });
  });

  it('compiles ls to a read intent', () => {
    const compiled = compileCommand(
      { verb: 'ls', path: '/doc' },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(compiled).toEqual({ kind: 'read', intent: { read: 'ls', path: '/doc' } });
  });

  it('compiles find to a read intent', () => {
    const ctx = { surface: 'excel' as const, mintChangeId: () => asChangeId('c1') };
    expect(compileCommand({ verb: 'find', path: '/work' }, ctx)).toEqual({
      kind: 'read',
      intent: { read: 'find', path: '/work' },
    });
    expect(compileCommand({ verb: 'find', path: '/work', glob: '*.tsv' }, ctx)).toEqual({
      kind: 'read',
      intent: { read: 'find', path: '/work', glob: '*.tsv' },
    });
  });

  it("compiles tail to a read intent (file-level DocFs verb, not compose.ts's pipeline tail transform)", () => {
    const ctx = { surface: 'excel' as const, mintChangeId: () => asChangeId('c1') };
    expect(compileCommand({ verb: 'tail', path: '/work/notes.md' }, ctx)).toEqual({
      kind: 'read',
      intent: { read: 'tail', path: '/work/notes.md' },
    });
    expect(compileCommand({ verb: 'tail', path: '/work/notes.md', n: 20 }, ctx)).toEqual({
      kind: 'read',
      intent: { read: 'tail', path: '/work/notes.md', n: 20 },
    });
  });

  it('compiles workspace verbs to local workspace intents', () => {
    expect(compileCommand({ verb: 'workspace' }, { surface: 'excel', mintChangeId: mint })).toEqual(
      {
        kind: 'workspace',
        intent: { workspace: 'list' },
      },
    );
    expect(
      compileCommand(
        { verb: 'save', name: 'schedule.tsv', source: { src: 'read', selector: 'A1:B9' } },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({
      kind: 'workspace',
      intent: {
        workspace: 'save',
        name: 'schedule.tsv',
        source: { src: 'read', selector: 'A1:B9' },
      },
    });
    expect(
      compileCommand(
        { verb: 'share', name: 'schedule.tsv', source: { src: 'read', selector: 'A1:B9' } },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({
      kind: 'workspace',
      intent: {
        workspace: 'share',
        name: 'schedule.tsv',
        source: { src: 'read', selector: 'A1:B9' },
      },
    });
    expect(
      compileCommand(
        { verb: 'cat', ref: 'schedule.tsv', head: 8 },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'workspace', intent: { workspace: 'cat', ref: 'schedule.tsv', head: 8 } });
    expect(
      compileCommand(
        { verb: 'grep', ref: 'schedule.tsv', pattern: 'Deep Work', context: 1 },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({
      kind: 'workspace',
      intent: { workspace: 'grep', ref: 'schedule.tsv', pattern: 'Deep Work', context: 1 },
    });
    expect(
      compileCommand(
        { verb: 'cp', src: 'a.tsv', dst: 'b.tsv' },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'workspace', intent: { workspace: 'cp', src: 'a.tsv', dst: 'b.tsv' } });
    expect(
      compileCommand(
        { verb: 'mv', src: 'a.tsv', dst: 'b.tsv' },
        { surface: 'excel', mintChangeId: mint },
      ),
    ).toEqual({ kind: 'workspace', intent: { workspace: 'mv', src: 'a.tsv', dst: 'b.tsv' } });
    expect(
      compileCommand({ verb: 'rm', name: 'a.tsv' }, { surface: 'excel', mintChangeId: mint }),
    ).toEqual({ kind: 'workspace', intent: { workspace: 'rm', name: 'a.tsv' } });
  });

  it('compiles control verbs', () => {
    expect(compileCommand({ verb: 'done' }, { surface: 'excel', mintChangeId: mint })).toEqual({
      kind: 'control',
      verb: 'done',
    });
    expect(
      compileCommand(
        { verb: 'help', topic: 'shape' },
        { surface: 'powerpoint', mintChangeId: mint },
      ),
    ).toEqual({
      kind: 'control',
      verb: 'help',
      topic: 'shape',
    });
  });

  /* ADR-0006 CLI parity verbs — each compiles to the kind + param shape its bridge consumes. */
  it('compiles `slide` → insert-slide with params.slide { title, bullets }', () => {
    const c = compileCommand(
      { verb: 'slide', title: 'Q3 Results', bullets: ['Revenue up 12%', 'Churn down'] },
      { surface: 'powerpoint', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-slide',
        surface: 'powerpoint',
        params: { slide: { title: 'Q3 Results', bullets: ['Revenue up 12%', 'Churn down'] } },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `/insert-slide deckBase64=…` into an explicit PowerPoint deck artifact import', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'insert-slide',
        props: {
          deckBase64: 'UEsDBBQ=',
          slideCount: '3',
          formatting: 'UseDestinationTheme',
          targetSlideId: '256#2',
          specFingerprint: '3a7c10ff',
        },
        args: [],
      },
      { surface: 'powerpoint', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-slide',
        surface: 'powerpoint',
        params: {
          deck: {
            base64: 'UEsDBBQ=',
            format: 'pptx',
            slideCount: 3,
            formatting: 'UseDestinationTheme',
            targetSlideId: '256#2',
            specFingerprint: '3a7c10ff',
          },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `shape` → set-shape-text with explicit slide + shape target', () => {
    const c = compileCommand(
      { verb: 'shape', selector: 'pp:shape:s2:s2-shape-1', text: 'Updated outlook' },
      { surface: 'powerpoint', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'set-shape-text',
        surface: 'powerpoint',
        params: {
          target: { slideId: 's2', shapeId: 's2-shape-1' },
          text: 'Updated outlook',
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('rejects `shape` selectors that do not include a slide id', () => {
    const c = compileCommand(
      { verb: 'shape', selector: 's2-shape-1', text: 'Updated outlook' },
      { surface: 'powerpoint', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      error: expect.stringContaining('shape selector must include slide and shape id'),
    });
  });

  it('compiles specialized /insert-text with a Word content anchor', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'insert-text',
        props: { text: ' Effective July 1.', match: 'This agreement begins', hint: 'Section 2' },
        args: [],
      },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-text',
        surface: 'word',
        params: {
          text: ' Effective July 1.',
          target: { matchText: 'This agreement begins', contextHint: 'Section 2' },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /fill-content-control from id/text props', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'fill-content-control',
        props: { id: 'CustomerName', text: 'VanArsdel, Ltd.' },
        args: [],
      },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'fill-content-control',
        surface: 'word',
        params: {
          text: 'VanArsdel, Ltd.',
          target: { contentControlId: 'CustomerName' },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `page` → append-page with target.matchText (title) + text (body)', () => {
    const c = compileCommand(
      { verb: 'page', title: 'Meeting notes', body: 'Decisions: ship' },
      { surface: 'onenote', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'append-page',
        surface: 'onenote',
        params: { target: { matchText: 'Meeting notes' }, text: 'Decisions: ship' },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `mail` → reply-mail with params.mail.body', () => {
    const c = compileCommand(
      { verb: 'mail', body: 'Thanks — confirming the dates below.' },
      { surface: 'outlook', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'reply-mail',
        surface: 'outlook',
        params: { mail: { body: 'Thanks — confirming the dates below.' } },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `post` → post-message with params.text', () => {
    const c = compileCommand(
      { verb: 'post', text: 'Summary of decisions: ...' },
      { surface: 'teams', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'post-message',
        surface: 'teams',
        params: { text: 'Summary of decisions: ...' },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `compose` → create-mail with params.mail.{subject,body}', () => {
    const c = compileCommand(
      { verb: 'compose', subject: 'Follow-up on Q3', body: 'Hi — summary below.' },
      { surface: 'outlook', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'create-mail',
        surface: 'outlook',
        params: { mail: { subject: 'Follow-up on Q3', body: 'Hi — summary below.' } },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `table` → create-table with params.table (ADR-0007)', () => {
    const c = compileCommand(
      { verb: 'table', range: 'Report!A1:C12', props: { headers: 'true', name: 'Top' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'create-table',
        surface: 'excel',
        params: { table: { range: 'Report!A1:C12', hasHeaders: true, name: 'Top' } },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `chart` → insert-chart with the typed chart payload (ADR-0007)', () => {
    const c = compileCommand(
      {
        verb: 'chart',
        chartType: 'column',
        range: 'Report!A1:B11',
        props: { title: 'Top regions', series: 'columns' },
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-chart',
        surface: 'excel',
        params: {
          chart: {
            chartType: 'column',
            sourceRange: 'Report!A1:B11',
            seriesBy: 'columns',
            title: 'Top regions',
          },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('rejects an unknown chart type at schema validation (corrective, not a bad write)', () => {
    const c = compileCommand(
      { verb: 'chart', chartType: 'donut', range: 'A1:B2', props: {} },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({ error: expect.stringContaining('insert-chart') });
  });

  it('compiles `cf` inline operator → format-conditional cellValue rule (ADR-0007)', () => {
    const c = compileCommand(
      { verb: 'cf', range: 'Sales!E2:E200', props: { op: '>', value: '100000', fill: '#C6EFCE' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'format-conditional',
        params: {
          conditional: {
            range: 'Sales!E2:E200',
            rule: { kind: 'cellValue', operator: 'gt', value: '100000', fill: '#C6EFCE' },
          },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `cf databar` and `cf top=N` into their typed rules', () => {
    const bar = compileCommand(
      { verb: 'cf', range: 'A:A', props: { databar: 'true' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(bar).toMatchObject({
      request: { params: { conditional: { rule: { kind: 'dataBar' } } } },
    });
    const top = compileCommand(
      { verb: 'cf', range: 'A:A', props: { top: '5', bottom: 'true' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(top).toMatchObject({
      request: { params: { conditional: { rule: { kind: 'top', rank: 5, bottom: true } } } },
    });
  });

  it('cf with no expressible rule is a corrective error', () => {
    const c = compileCommand(
      { verb: 'cf', range: 'A:A', props: { nonsense: 'true' } },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({ error: expect.stringContaining('rule') });
  });

  it('compiles `spill` (resolved cells) → a write-cells grid (ADR-0007 §3)', () => {
    const c = compileCommand(
      {
        verb: 'spill',
        range: 'Report!A1',
        cells: [
          ['Region', 'Revenue'],
          ['EMEA', '120'],
        ],
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'write-cells',
        params: {
          target: { range: 'Report!A1' },
          cells: [
            ['Region', 'Revenue'],
            ['EMEA', '120'],
          ],
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles `grid` literal TSV → one write-cells request', () => {
    const c = compileCommand(
      {
        verb: 'grid',
        range: "'Daily schedule'!C5:D6",
        cells: [
          ['Monday', 'Tuesday'],
          ['Deep Work', 'Music Lesson'],
        ],
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'write-cells',
        params: {
          target: { range: "'Daily schedule'!C5:D6" },
          cells: [
            ['Monday', 'Tuesday'],
            ['Deep Work', 'Music Lesson'],
          ],
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /insert-pivot into typed Excel PivotTable params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'insert-pivot',
        props: {
          sourceRange: 'Sales!A1:G500',
          destinationRange: 'Pivot!A3',
          rowFields: 'Region,Segment',
          valueFields: 'Revenue',
          name: 'RevenuePivot',
        },
        args: [],
      },
      { surface: 'excel', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-pivot',
        surface: 'excel',
        params: {
          pivot: {
            sourceRange: 'Sales!A1:G500',
            destinationRange: 'Pivot!A3',
            rowFields: ['Region', 'Segment'],
            valueFields: ['Revenue'],
            name: 'RevenuePivot',
          },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /format-shape into typed PowerPoint shape formatting params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'format-shape',
        props: {
          ref: 'pp:shape:s2:shape7',
          fill: '#0F6CBD',
          fontBold: 'true',
          fontSize: '18',
        },
        args: [],
      },
      { surface: 'powerpoint', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'format-shape',
        surface: 'powerpoint',
        params: {
          target: { slideId: 's2', shapeId: 'shape7' },
          shapeFormat: { fill: '#0F6CBD', font: { bold: true, size: 18 } },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /insert-content-control into typed Word params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'insert-content-control',
        props: {
          match: 'Customer:',
          type: 'richText',
          tag: 'CustomerName',
          title: 'Customer name',
        },
        args: [],
      },
      { surface: 'word', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'insert-content-control',
        surface: 'word',
        params: {
          target: { matchText: 'Customer:' },
          contentControl: { type: 'richText', tag: 'CustomerName', title: 'Customer name' },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /set-recipients into typed Outlook draft recipient params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'set-recipients',
        props: {
          to: 'a@example.com;b@example.com',
          cc: 'reviewer@example.com',
          recipientMode: 'add',
        },
        args: [],
      },
      { surface: 'outlook', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'set-recipients',
        surface: 'outlook',
        params: {
          mail: {
            to: ['a@example.com', 'b@example.com'],
            cc: ['reviewer@example.com'],
            bcc: [],
            recipientMode: 'add',
          },
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles specialized /add-outline into typed OneNote content params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'add-outline',
        props: { html: '<p>Action items</p>' },
        args: [],
      },
      { surface: 'onenote', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'add-outline',
        surface: 'onenote',
        params: { html: '<p>Action items</p>' },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });

  it('compiles estate-gated /post-channel-message into typed Teams Graph target params', () => {
    const c = compileCommand(
      {
        verb: 'invoke',
        kind: 'post-channel-message',
        props: {
          teamId: 'team-1',
          channelId: 'channel-1',
          text: 'Standup notes ready',
        },
        args: [],
      },
      { surface: 'teams', mintChangeId: mint },
    );
    expect(c).toMatchObject({
      kind: 'write',
      request: {
        kind: 'post-channel-message',
        surface: 'teams',
        params: {
          graphTarget: { teamId: 'team-1', channelId: 'channel-1' },
          text: 'Standup notes ready',
        },
      },
    });
    if ('request' in c) expect(() => ActuationRequestSchema.parse(c.request)).not.toThrow();
  });
});

describe('renderGrammarPrompt', () => {
  it('advertises set for Excel and not suggest', () => {
    const prompt = renderGrammarPrompt(excelManifest);
    expect(prompt).toContain('set <A1> <value|=formula>');
    expect(prompt).toContain('grid <range> = "a\\tb\\nc\\td"');
    expect(prompt).toContain('read <A1|NamedRange>');
    expect(prompt).toContain('context [incremental|inline-preferred');
    expect(prompt).not.toContain('suggest "old text"');
    expect(prompt).toContain('```cmd');
  });

  it('advertises chart only when insert-chart is present', () => {
    expect(renderGrammarPrompt(excelManifest)).not.toContain('chart <column|bar|line|pie');
    const prompt = renderGrammarPrompt({
      ...excelManifest,
      actuations: [
        ...excelManifest.actuations,
        {
          kind: 'insert-chart',
          surface: 'excel',
          title: 'Insert chart',
          reversible: true,
        },
      ],
    });
    expect(prompt).toContain('chart <column|bar|line|pie|scatter|area>');
  });

  it('advertises suggest for Word and not set', () => {
    const prompt = renderGrammarPrompt(wordManifest);
    expect(prompt).toContain('suggest "old text" => "new text"');
    expect(prompt).not.toContain('set <A1');
  });

  it('advertises specialized slash operations only when the surface manifest includes them', () => {
    expect(renderGrammarPrompt(wordManifest)).not.toContain('/fill-content-control');
    const prompt = renderGrammarPrompt({
      ...wordManifest,
      actuations: [
        ...wordManifest.actuations,
        { kind: 'insert-text', surface: 'word', title: 'Insert text', reversible: false },
        {
          kind: 'fill-content-control',
          surface: 'word',
          title: 'Fill content control',
          reversible: true,
        },
      ],
    });
    expect(prompt).toContain('/insert-text [key=value ...]');
    expect(prompt).toContain('/fill-content-control [key=value ...]');
  });

  it('uses registry metadata for advanced slash operations once advertised', () => {
    const prompt = renderGrammarPrompt({
      ...excelManifest,
      actuations: [
        ...excelManifest.actuations,
        {
          kind: 'insert-pivot',
          surface: 'excel',
          title: 'Insert PivotTable',
          reversible: true,
        },
      ],
    });
    expect(prompt).toContain('/insert-pivot [key=value ...]');
    expect(prompt).toContain('summarize a table/range');
  });

  it('explicitly rejects non-cmd fences and thinking prose', () => {
    const prompt = renderGrammarPrompt(excelManifest);
    expect(prompt).toContain('Never emit prose, thinking');
    expect(prompt).toContain('```python');
    expect(prompt).toContain('invalid and will be ignored');
  });
});

describe('renderCommandHelp', () => {
  it('renders targeted generated help for an available command', () => {
    const prompt = renderCommandHelp(
      {
        surface: 'powerpoint',
        contextKinds: ['slide', 'shape'],
        reads: ['outline', 'read'],
        actuations: [
          { kind: 'insert-slide', surface: 'powerpoint', title: 'Insert slide', reversible: true },
          {
            kind: 'set-shape-text',
            surface: 'powerpoint',
            title: 'Replace shape text',
            reversible: true,
          },
        ],
      },
      'shape',
    );

    expect(prompt).toContain('Command: shape');
    expect(prompt).toContain('Discovery sequence');
    expect(prompt).toContain('shape <pp:shape:slideId:shapeId> "new text"');
  });

  it('fails closed for targeted write help when the capability is unavailable', () => {
    const prompt = renderCommandHelp(wordManifest, 'shape');
    expect(prompt).toContain('Command unavailable on this surface: shape');
    expect(prompt).toContain('does not advertise set-shape-text');
  });

  it('renders targeted help for an available specialized slash operation', () => {
    const prompt = renderCommandHelp(
      {
        ...wordManifest,
        actuations: [
          ...wordManifest.actuations,
          {
            kind: 'fill-content-control',
            surface: 'word',
            title: 'Fill content control',
            reversible: true,
          },
        ],
      },
      '/fill-content-control',
    );
    expect(prompt).toContain('Command: /fill-content-control');
    expect(prompt).toContain('Discovery sequence');
    expect(prompt).toContain('id=<contentControlId>');
  });

  it('renders registry-backed targeted help for an available advanced slash operation', () => {
    const prompt = renderCommandHelp(
      {
        ...excelManifest,
        actuations: [
          ...excelManifest.actuations,
          {
            kind: 'insert-pivot',
            surface: 'excel',
            title: 'Insert PivotTable',
            reversible: true,
          },
        ],
      },
      '/insert-pivot',
    );
    expect(prompt).toContain('Command: /insert-pivot');
    expect(prompt).toContain('Excel native PivotTable');
    expect(prompt).toContain('Registry status: promotable');
    expect(prompt).toContain('Preview must show: source range');
  });
});

/* ───────────────────────── loop fixtures ──────────────────────────────── */

const excelManifest: CapabilityManifest = {
  surface: 'excel',
  contextKinds: ['range', 'sheet'],
  reads: ['outline', 'read', 'search'],
  actuations: [{ kind: 'write-cells', surface: 'excel', title: 'Write cells', reversible: true }],
};

const wordManifest: CapabilityManifest = {
  surface: 'word',
  contextKinds: ['selection', 'document'],
  reads: ['outline', 'read', 'search'],
  actuations: [
    { kind: 'tracked-change', surface: 'word', title: 'Insert tracked change', reversible: true },
  ],
};

function snapshot(surface: 'excel' | 'word', version: number): DocStateSnapshot {
  return {
    surface,
    version,
    capturedAt: '2026-06-22T00:00:00Z',
    outline: [],
    inventory: [],
  };
}

/** A fake Excel bridge recording actuations, reads, and serving a versioned doc-state. */
class FakeExcelBridge implements DocBridge {
  readonly surface = 'excel' as const;
  applied: ActuationRequest[] = [];
  reads: string[] = [];
  revealed: ContextRef[] = [];
  version = 1;
  contextRefs: ContextRef[] = [
    {
      id: 'xl:Sales!A1:C9',
      kind: 'range',
      surface: 'excel',
      title: 'Sales!A1:C9',
      preview: 'Region | Revenue',
      hostRef: { type: 'excel.range', worksheet: 'Sales', address: 'A1:C9' },
      live: true,
    },
  ];

  getCapabilities(): CapabilityManifest {
    return excelManifest;
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve(this.contextRefs);
  }
  resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    return Promise.resolve([
      {
        ref,
        value: { as: 'text', text: `values of ${ref.title}` },
      },
    ]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    this.version += 1;
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'F2',
    });
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    return Promise.resolve(snapshot('excel', this.version));
  }
  readRange(a1: string): Promise<ResolvedContext[]> {
    this.reads.push(a1);
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: { as: 'text', text: `values of ${a1}` },
      },
    ]);
  }
  searchDocument(query: string): Promise<ResolvedContext[]> {
    this.reads.push(`search:${query}`);
    return Promise.resolve([
      {
        ref: { id: `xl:search`, kind: 'range', surface: 'excel', title: query, live: false },
        value: { as: 'text', text: `rows matching ${query}` },
      },
    ]);
  }
  canRevealContext(ref: ContextRef): boolean {
    return ref.surface === 'excel' && ref.kind === 'range';
  }
  revealContext(ref: ContextRef): Promise<void> {
    this.revealed.push(ref);
    return Promise.resolve();
  }
}

class FakeWordBridge implements DocBridge {
  readonly surface = 'word' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return wordManifest;
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({
      ok: true,
      changeId: request.changeId,
      kind: request.kind,
      location: 'para:1',
    });
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    return Promise.resolve(snapshot('word', 1));
  }
  searchDocument(): Promise<ResolvedContext[]> {
    return Promise.resolve([
      {
        ref: { id: 'w:1', kind: 'paragraph', surface: 'word', title: 'p1', live: false },
        value: { as: 'text', text: 'The SLA is 99.5%.' },
      },
    ]);
  }
}

/**
 * A fake StreamAssistClient that replays a scripted transcript: one string of answer text per
 * model turn (wrapped in token + provenance + done SSE events), or a prebuilt SSE event list for
 * lower-level protocol branches. Records the queries it received.
 */
function fakeClient(turns: Array<string | SseEvent[]>): {
  client: StreamAssistClient;
  queries: string[];
} {
  const queries: string[] = [];
  let i = 0;
  const stream = async function* (
    req: AssistRequest,
    _opts: StreamOptions,
  ): AsyncGenerator<SseEvent> {
    queries.push(req.query ?? '');
    const scripted = turns[i++] ?? '```cmd\ndone\n```';
    if (typeof scripted === 'string') {
      yield { type: 'token', text: scripted };
    } else {
      for (const event of scripted) yield event;
    }
    yield {
      type: 'provenance',
      payload: {
        agentId: 'gemini-enterprise:e',
        identity: 'v.k@acme',
        timestamp: '2026-06-22T00:00:00Z',
        sources: [],
        contentHash: 'h',
        sessionId: 'sess_loop' as never,
      },
    };
    yield { type: 'done' };
  };
  const client = { stream } as unknown as StreamAssistClient;
  return { client, queries };
}

const unit = { connectors: [], surfaceContext: { kind: 'excel' as const } };

async function collect(
  gen: AsyncGenerator<SseEvent | CommandLoopEvent>,
): Promise<Array<SseEvent | CommandLoopEvent>> {
  const out: Array<SseEvent | CommandLoopEvent> = [];
  for await (const e of gen) out.push(e);
  return out;
}

function loopEvents(events: Array<SseEvent | CommandLoopEvent>): CommandLoopEvent[] {
  const kinds = new Set([
    'turn-start',
    'command',
    'read-result',
    'write-result',
    'no-fence',
    'capped',
    'done',
    'exhausted',
  ]);
  return events.filter((e) => kinds.has(e.type)) as CommandLoopEvent[];
}

/* ───────────────────────── the loop ───────────────────────────────────── */

describe('AssistSession.runCommands — the bounded command loop', () => {
  it('read-many: batches all reads in a turn, then terminates on done', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '**thought** discover first\n```cmd\noutline\nread Sales!C2:C7\nsearch margin\n```',
      '**answer** all set\n```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Analyze the sheet'));
    const loop = loopEvents(events);

    // Three reads executed in turn 1.
    const reads = loop.filter((e) => e.type === 'read-result');
    expect(reads).toHaveLength(3);
    expect(bridge.reads).toEqual(['Sales!C2:C7', 'search:margin']); // outline uses captureDocState
    // Terminated on done.
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
    // Turn 1 query carries the protocol + the task.
    expect(queries[0]).toContain('TASK:');
    expect(queries[0]).toContain('Analyze the sheet');
    // Turn 2 query is the ```result block fed back.
    expect(queries[1]).toContain('```result');
  });

  it('ls and find DocFs verbs are dispatched as reads through the full command loop', async () => {
    // Regression guard: isReadCommand() gates dispatch on a hand-maintained READ_COMMAND_VERBS set
    // in assist-session.ts that is separate from command-grammar.ts's READ_VERBS — a verb can parse
    // and compile correctly yet never reach runReadIntent if it is missing from that set (exactly what
    // happened for `ls` when it first landed: it worked in isolation but was routed as an unsupported
    // effect verb inside runCommands()). Exercise both `ls` and `find` end-to-end, not just compile.
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['```cmd\nls /doc\nfind /doc\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('list DocFs entries'));
    const loop = loopEvents(events);

    const reads = loop.filter((e) => e.type === 'read-result');
    expect(reads).toHaveLength(2);
    for (const read of reads) {
      expect(read).toMatchObject({
        result: expect.not.objectContaining({ error: expect.anything() }),
      });
    }
    expect(loop.some((e) => e.type === 'command' && 'error' in e.compiled)).toBe(false);
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
  });

  it('tail (file-level DocFs verb) is dispatched as a read through the full command loop', async () => {
    // Same regression guard as the `ls`/`find` test above, for `tail`: READ_COMMAND_VERBS in
    // assist-session.ts is a THIRD hand-maintained verb set (separate from command-grammar.ts's
    // READ_VERBS and command-protocol.ts's compileCommand switch) that gates whether a read verb
    // is actually reachable through runCommands(). `ls`/`find` both parsed and compiled correctly
    // in isolation yet were unreachable here until added to that set — exercise `tail` end-to-end,
    // not just compileCommand/runReadIntent in isolation. `/doc/outline.md` always exists (it is
    // the doc-mount's outline view, even when the outline is empty), so no extra fixture is needed.
    // This is the DocFs file-level `tail`, distinct from compose.ts's pipeline `tail` transform
    // (`(... | tail 5)`, a different grammar slot entirely).
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['```cmd\ntail /doc/outline.md\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('show the tail of the outline'));
    const loop = loopEvents(events);

    const reads = loop.filter((e) => e.type === 'read-result');
    expect(reads).toHaveLength(1);
    expect(reads[0]).toMatchObject({
      type: 'read-result',
      intentLabel: 'tail /doc/outline.md',
      result: [{ text: '' }], // the fake bridge's doc-state has an empty outline.
    });
    expect(loop.some((e) => e.type === 'command' && 'error' in e.compiled)).toBe(false);
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
  });

  it('context returns upload and code-execution strategy without reading or writing host content', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '```cmd\ncontext analytical full-scope upload-preferred code-execution-preferred\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Analyze the full workbook'));
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      intentLabel: 'context analytical full-scope upload-preferred code-execution-preferred',
      result: {
        strategy: {
          scope: 'whole-artifact',
          transfer: 'upload-candidate',
          analysis: 'code-execution-candidate',
        },
        upload: {
          state: 'recommended',
          supportedFormats: expect.arrayContaining([
            {
              extension: '.xlsx',
              mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            },
          ]),
        },
      },
    });
    expect(bridge.reads).toEqual([]);
    expect(bridge.applied).toEqual([]);
    expect(queries[1]).toContain('"upload"');
    expect(queries[1]).toContain('fileId');
  });

  it('workspace save stores a host read once and grep searches the artifact locally', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient([
      '```cmd\nsave schedule.txt = read Sales!C2:C7\ngrep schedule.txt "Sales" context=0\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Stage data locally'));
    const reads = loopEvents(events).filter((e) => e.type === 'read-result');

    expect(bridge.reads).toEqual(['Sales!C2:C7']);
    expect(reads[0]).toMatchObject({
      type: 'read-result',
      intentLabel: 'save schedule.txt',
      result: { workspace: 'save', artifact: { name: 'schedule.txt', id: 'ws:1' } },
    });
    expect(reads[1]).toMatchObject({
      type: 'read-result',
      intentLabel: 'grep schedule.txt',
      result: {
        workspace: 'grep',
        artifact: { name: 'schedule.txt' },
        pattern: 'Sales',
        matches: [{ line: 1, text: 'values of Sales!C2:C7' }],
      },
    });
  });

  function fakeSharedStore() {
    const shared = new Map<string, string>();
    return {
      shared,
      sharedStore: {
        list: () =>
          Promise.resolve(
            [...shared.entries()].map(([name, text]) => ({ name, size: text.length })),
          ),
        read: (path: string) => Promise.resolve(shared.get(path)),
        write: (path: string, content: string) => {
          shared.set(path, content);
          return Promise.resolve();
        },
        remove: (path: string) => {
          shared.delete(path);
          return Promise.resolve();
        },
      },
    };
  }

  it('workspace share is blocked by default — estate writes are disabled fail-closed', async () => {
    const bridge = new FakeExcelBridge();
    const { sharedStore } = fakeSharedStore();
    const { client } = fakeClient([
      '```cmd\nshare schedule.txt = read Sales!C2:C7\n```',
      '```cmd\ndone\n```',
    ]);
    // No `estateWritesEnabled` — even with a configured sharedStore, share must stay inert.
    const session = new AssistSession(bridge, client, { unit, sharedStore });

    const events = await collect(session.runCommands('Publish data cross-surface'));
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      result: { workspace: 'error', error: expect.stringContaining('estate writes are disabled') },
    });
  });

  it('workspace share with estate writes enabled but no sharedStore returns a corrective error', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient([
      '```cmd\nshare schedule.txt = read Sales!C2:C7\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit, estateWritesEnabled: true });

    const events = await collect(session.runCommands('Publish data cross-surface'));
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      result: { workspace: 'error', error: expect.stringContaining('sharing is not configured') },
    });
  });

  it('workspace share with no approveShare supplied is blocked fail-closed (no silent write)', async () => {
    const bridge = new FakeExcelBridge();
    const { sharedStore, shared } = fakeSharedStore();
    const { client } = fakeClient([
      '```cmd\nshare schedule.txt = read Sales!C2:C7\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, {
      unit,
      sharedStore,
      estateWritesEnabled: true,
    });

    // runCommands() called WITHOUT approveShare in opts.
    const events = await collect(session.runCommands('Publish data cross-surface'));
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      result: { workspace: 'error', error: expect.stringContaining('requires user approval') },
    });
    expect(shared.size).toBe(0);
  });

  it('workspace share writes to the sharedStore only after estateWritesEnabled + approveShare both grant it, and stamps a provenance companion file', async () => {
    const bridge = new FakeExcelBridge();
    const { sharedStore, shared } = fakeSharedStore();
    const { client } = fakeClient([
      '```cmd\nshare schedule.txt = read Sales!C2:C7\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, {
      unit,
      sharedStore,
      estateWritesEnabled: true,
    });

    const events = await collect(
      session.runCommands('Publish data cross-surface', { approveShare: () => true }),
    );
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(bridge.reads).toEqual(['Sales!C2:C7']);
    expect(read).toMatchObject({
      type: 'read-result',
      intentLabel: 'share schedule.txt',
      result: { workspace: 'share', name: 'schedule.txt' },
    });
    expect(shared.get('schedule.txt')).toContain('values of Sales!C2:C7');
    const provenance = JSON.parse(shared.get('schedule.txt.provenance.json') ?? '{}');
    expect(provenance).toMatchObject({ agentId: 'gemini-enterprise:e', identity: 'v.k@acme' });
  });

  it('workspace share is blocked when approveShare explicitly denies it', async () => {
    const bridge = new FakeExcelBridge();
    const { sharedStore, shared } = fakeSharedStore();
    const { client } = fakeClient([
      '```cmd\nshare schedule.txt = read Sales!C2:C7\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, {
      unit,
      sharedStore,
      estateWritesEnabled: true,
    });

    const events = await collect(
      session.runCommands('Publish data cross-surface', { approveShare: () => false }),
    );
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      result: { workspace: 'error', error: expect.stringContaining('requires user approval') },
    });
    expect(shared.size).toBe(0);
  });

  it('lists and inspects addressable context without creating a write plan', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '```cmd\nlist range\nproperties xl:Sales!A1:C9\ninspect xl:Sales!A1:C9\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Inspect the active range'));
    const loop = loopEvents(events);

    expect(loop.filter((e) => e.type === 'read-result')).toHaveLength(3);
    expect(loop.some((e) => e.type === 'plan-preview')).toBe(false);
    expect(bridge.applied).toEqual([]);
    expect(queries[1]).toContain('hostRef');
    expect(queries[1]).toContain('values of Sales!A1:C9');
  });

  it('open navigates to a revealable context ref and does not mutate the workbook', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '```cmd\nopen xl:Sales!A1:C9\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('Open the Sales range'));
    const read = loopEvents(events).find((e) => e.type === 'read-result');

    expect(read).toMatchObject({
      type: 'read-result',
      intentLabel: 'open xl:Sales!A1:C9',
      result: { opened: true, navigationOnly: true },
    });
    expect(bridge.revealed.map((r) => r.id)).toEqual(['xl:Sales!A1:C9']);
    expect(bridge.applied).toEqual([]);
    expect(queries[1]).toContain('"navigationOnly":true');
  });

  it('write-one: a set compiles to a gated write-cells request, one at a time', async () => {
    const bridge = new FakeExcelBridge();
    const gate = new TriggerRegistry();
    const seen: ActuationRequest[] = [];
    gate.register({
      id: 'audit',
      on: 'pre-actuation',
      handle: (e) => {
        if (e.type === 'pre-actuation') seen.push(e.request);
        return { kind: 'continue' };
      },
    });
    const { client } = fakeClient([
      '```cmd\nset Sales!F2 =SUM(C2:C7)\nset Sales!F3 =SUM(D2:D7)\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit, triggers: gate });

    const events = await collect(session.runCommands('Write totals', { approveWrite: () => true }));
    const writes = loopEvents(events).filter((e) => e.type === 'write-result');

    expect(writes).toHaveLength(2);
    expect(bridge.applied).toHaveLength(2);
    expect(seen).toHaveLength(2); // each gated
    expect(bridge.applied[0]).toMatchObject({
      kind: 'write-cells',
      surface: 'excel',
      params: { target: { range: 'Sales!F2' }, cells: [['=SUM(C2:C7)']] },
    });
    // Provenance stamped from the streamed turn.
    expect(bridge.applied[0]!.provenance?.identity).toBe('v.k@acme');
  });

  it('a blocked gate yields a corrective write-result, not a thrown loop', async () => {
    const bridge = new FakeExcelBridge();
    const gate = new TriggerRegistry();
    gate.register({
      id: 'veto',
      on: 'pre-actuation',
      handle: () => ({ kind: 'block', reason: 'needs approval' }),
    });
    const { client } = fakeClient(['```cmd\nset A1 5\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit, triggers: gate });

    const events = await collect(session.runCommands('write', { approveWrite: () => true }));
    const write = loopEvents(events).find((e) => e.type === 'write-result');
    expect(write).toMatchObject({ result: { ok: false, error: { code: 'blocked' } } });
    expect(bridge.applied).toHaveLength(0); // never actuated
  });

  it('fail-closed: a write with no approver is refused and never actuated', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['```cmd\nset A1 5\n```', '```cmd\ndone\n```']);
    const session = new AssistSession(bridge, client, { unit });
    // No approvePlan AND no approveWrite → the ADR-0005 Phase-2 plan is blocked fail-closed (the
    // DocBridge confirmation contract): the effect is refused and never actuated.
    const events = await collect(session.runCommands('write'));
    const write = loopEvents(events).find((e) => e.type === 'write-result');
    expect(write).toMatchObject({ result: { ok: false, error: { code: 'plan_unapproved' } } });
    expect(bridge.applied).toHaveLength(0);
  });

  it('write-one cap: only maxWritesPerTurn writes actuate in one block', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient([
      '```cmd\nset A1 1\nset A2 2\nset A3 3\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });
    const events = await collect(
      session.runCommands('write three', { approveWrite: () => true, maxWritesPerTurn: 2 }),
    );
    expect(bridge.applied).toHaveLength(2); // the third write is capped, not actuated
    expect(loopEvents(events).some((e) => e.type === 'capped')).toBe(true);
  });

  it('suggest → tracked-change with target.matchText (Word)', async () => {
    const bridge = new FakeWordBridge();
    const wordUnit = { connectors: [], surfaceContext: { kind: 'word' as const } };
    const { client } = fakeClient([
      '```cmd\nsuggest "The SLA is 99.5%." => "The SLA is ~99.5% (source needed)."\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit: wordUnit });

    await collect(session.runCommands('flag unsourced claims', { approveWrite: () => true }));
    expect(bridge.applied[0]).toMatchObject({
      kind: 'tracked-change',
      params: {
        target: { matchText: 'The SLA is 99.5%.' },
        text: 'The SLA is ~99.5% (source needed).',
      },
    });
  });

  it('an unknown verb returns a corrective error the next turn self-corrects', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '```cmd\nsett A1 5\n```', // typo
      '```cmd\nset A1 5\n```', // corrected
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    await collect(session.runCommands('write a cell', { approveWrite: () => true }));

    // The corrective error was fed back on turn 2's query.
    expect(queries[1]).toContain('unknown verb');
    expect(queries[1]).toContain('did you mean');
    // The self-corrected write landed.
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]).toMatchObject({ params: { target: { range: 'A1' } } });
  });

  it('a no-fence turn re-prompts once, then proceeds', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      '**thought** I am still thinking, no commands yet.', // no fence
      '```cmd\nset A1 1\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('write a cell', { approveWrite: () => true }));
    const loop = loopEvents(events);

    expect(loop.some((e) => e.type === 'no-fence')).toBe(true);
    // The re-prompt query nudges for a cmd block.
    expect(queries[1]).toContain('```cmd');
    expect(queries[1]).toContain('Do not emit prose');
    expect(queries[1]).toContain('```python');
    expect(bridge.applied).toHaveLength(1); // still completed the write
    expect(loop.at(-1)).toMatchObject({ type: 'done' });
  });

  it('reprompts hosted code-execution output back into the Office CLI route', async () => {
    const bridge = new FakeExcelBridge();
    const { client, queries } = fakeClient([
      [
        { type: 'activity', text: 'Building chart data' },
        { type: 'code-execution', language: 'python', code: 'import matplotlib.pyplot as plt' },
        { type: 'code-execution-result', outcome: 'OUTCOME_OK', output: 'image/png;base64,...' },
      ],
      '```cmd\nset Sales!Z1 "chart command required, not hosted code"\n```',
      '```cmd\ndone\n```',
    ]);
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(
      session.runCommands('visualize the sales table', { approveWrite: () => true }),
    );
    const loop = loopEvents(events);

    expect(loop.some((e) => e.type === 'no-fence')).toBe(true);
    expect(queries[1]).toContain('Hosted Python/code execution is not a valid executor response');
    expect(queries[1]).toContain('emit the Office chart command');
    expect(bridge.applied).toHaveLength(1);
    expect(bridge.applied[0]).toMatchObject({
      kind: 'write-cells',
      params: { target: { range: 'Sales!Z1' } },
    });
  });

  it('stops at maxTurns without done (exhausted)', async () => {
    const bridge = new FakeExcelBridge();
    // Every turn emits a read, never done.
    const { client } = fakeClient(Array(20).fill('```cmd\noutline\n```'));
    const session = new AssistSession(bridge, client, { unit });

    const events = await collect(session.runCommands('loop forever', { maxTurns: 3 }));
    const loop = loopEvents(events);
    expect(loop.filter((e) => e.type === 'turn-start')).toHaveLength(3);
    expect(loop.at(-1)).toMatchObject({ type: 'exhausted', turns: 3 });
  });

  it('leaves plain ask() unchanged (still streams a grounded answer)', async () => {
    const bridge = new FakeExcelBridge();
    const { client } = fakeClient(['hello world']);
    const session = new AssistSession(bridge, client, { unit });
    const out: SseEvent[] = [];
    for await (const e of session.ask('hi')) out.push(e as SseEvent);
    expect(out.map((e) => e.type)).toContain('token');
  });
});
