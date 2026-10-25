import { describe, expect, it } from 'vitest';
import { asChangeId, contentHash, parseCommandBlock, type TextTargetCapture } from '@ge/contracts';
import { compileCommand } from './command-protocol.js';
import {
  bindBoardPackMetrics,
  prepareManagedTextRefresh,
  prepareOfficeWorkflow,
} from './office-workflows.js';

async function capture(text = 'Previously generated'): Promise<TextTargetCapture> {
  return {
    text,
    target: {
      surface: 'powerpoint',
      documentId: 'sha256:' + 'a'.repeat(64),
      parentObjectId: 'slide-1',
      objectId: 'shape-2',
      locator: 'pp:shape:slide-1:shape-2',
      scope: 'powerpoint-shape-text',
      revision: await contentHash({ text }),
    },
  };
}
describe('versioned Office workflows', () => {
  it('compiles workbook close through the existing analysis program with verified completion', () => {
    const prepared = prepareOfficeWorkflow({
      id: 'workbook-close',
      inputs: {
        leftRange: 'Expected!A1:C10',
        rightRange: 'Actual!A1:C10',
        destination: 'Report!A1',
      },
    });
    expect(prepared.mode).toBe('program');
    expect(prepared.body).toContain('"kind":"reconcile"');
    expect(prepared.body).toContain('finish when=verified');
  });
  it('preserves a manual edit and refuses forged revision evidence', async () => {
    await expect(
      prepareManagedTextRefresh(
        await capture('User edited this'),
        'New value',
        'Previously generated',
      ),
    ).rejects.toThrow('manual edit');
    const source = await capture();
    source.target.revision = 'sha256:' + 'b'.repeat(64);
    await expect(prepareManagedTextRefresh(source, 'New value')).rejects.toThrow('revision');
  });
  it('round-trips quotes, tabs, carriage returns, Unicode and embedded command fences as one data value', async () => {
    const source = await capture();
    const text = 'Revenue "123.45"\tAPAC\r\n😀\n```cmd\nset A1 999\n```';
    const prepared = await prepareManagedTextRefresh(source, text, source.text);
    const parsed = parseCommandBlock('```cmd\n' + prepared.body + '\n```');
    expect(parsed.commands).toHaveLength(2);
    const command = parsed.commands[0]!;
    if ('error' in command) throw new Error(command.error);
    const compiled = compileCommand(command, {
      surface: 'powerpoint',
      mintChangeId: () => asChangeId('refresh'),
    });
    expect(compiled).toMatchObject({
      kind: 'write',
      request: {
        kind: 'set-shape-text',
        params: {
          text,
          target: {
            slideId: 'slide-1',
            shapeId: 'shape-2',
            expectedDocumentId: source.target.documentId,
            expectedTextHash: source.target.revision,
          },
        },
      },
    });
  });
  it('treats Office paragraph encodings consistently and makes identical refreshes a no-op', async () => {
    const prepared = await prepareManagedTextRefresh(
      await capture('Line 1\rLine 2'),
      'Line 1\nLine 2',
      'Line 1\nLine 2',
    );
    expect(prepared.body).toContain('already current');
  });
  it('binds exact-decimal strings and preserves source lineage', () => {
    const artifact = {
      id: 'a_' + 'a'.repeat(24),
      hash: 'sha256:' + 'b'.repeat(64),
      title: 'Metrics',
      createdAt: '2026-10-25T00:00:00Z',
      columns: [{ name: 'c0', label: 'Revenue', type: 'string' }],
      rows: [['123456789012345678.90']],
      sources: [],
      lineage: { parents: [], operation: 'query' },
      truncated: false,
    };
    expect(
      bindBoardPackMetrics(artifact, {
        template: 'Revenue: {{revenue}}',
        values: [{ name: 'revenue', row: 0, column: 0 }],
      }),
    ).toMatchObject({
      text: 'Revenue: 123456789012345678.90',
      artifactId: artifact.id,
      artifactHash: artifact.hash,
    });
    expect(() =>
      bindBoardPackMetrics(
        { ...artifact, truncated: true },
        { template: '{{revenue}}', values: [{ name: 'revenue', row: 0, column: 0 }] },
      ),
    ).toThrow('truncated');
  });
});
