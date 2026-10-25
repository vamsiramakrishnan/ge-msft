import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ResolvedContext, SseEvent } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';
import { AssistSession } from '@ge/runtime';
import { OutlookBridge } from '@ge/bridge-outlook';
import { WordBridge, type WordHost } from '@ge/bridge-word';
import { installFakeOutlook, outlookSeed } from '../test-harness/fake-outlook.js';

afterEach(() => {
  vi.restoreAllMocks();
});

function recordingClient() {
  const contexts: ResolvedContext[][] = [];
  const client = {
    stream: async function* (_req: unknown, opts: StreamOptions): AsyncGenerator<SseEvent> {
      contexts.push(opts.context ?? []);
      yield { type: 'token', text: 'Answer from supplied evidence only.' };
      yield { type: 'done' };
    },
  } as unknown as StreamAssistClient;
  return { client, contexts };
}
const textOf = (parts: ResolvedContext[]): string =>
  parts.flatMap((part) => (part.value.as === 'text' ? [part.value.text] : [])).join('\n');

describe('bounded automatic chat reads through real bridges', () => {
  it('uses one Word page when probes miss and discloses the omitted paragraphs and next page', async () => {
    const paragraphs = Array.from({ length: 6000 }, (_, i) => ({
      text: `Paragraph ${i + 1}: the board discussed the operating plan.`,
      styleBuiltIn: 'Normal',
    }));
    const readParagraphs = vi.fn(async () => paragraphs);
    const readParagraphPage = vi.fn(async (offset: number, count: number) =>
      paragraphs.slice(offset, offset + count),
    );
    const host = {
      readParagraphs,
      readParagraphPage,
      searchText: vi.fn(async () => []),
    } as unknown as WordHost;
    const { client, contexts } = recordingClient();
    const session = new AssistSession(new WordBridge(host), client, {
      unit: { connectors: [], surfaceContext: { kind: 'word' } },
      context: { docState: false },
    });
    for await (const _event of session.ask('Apa keputusan rapat?')) void _event;
    expect(readParagraphs).not.toHaveBeenCalled();
    expect(readParagraphPage).toHaveBeenCalledTimes(1);
    expect(readParagraphPage).toHaveBeenCalledWith(0, 41);
    const text = textOf(contexts[0]!);
    expect(text).toContain('Paragraph 1:');
    expect(text).not.toContain('Paragraph 6000:');
    expect(text).toContain('"omittedParts":36');
    expect(text).toContain('"eof":false');
    expect(text).toContain('read document:40');
    expect(text).toContain('headers, footers');
    expect(session.context.size).toBe(0);
  });

  it('reports a failed Word page without retrying through its unbounded resolver', async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const readParagraphs = vi.fn(async () => []);
    const host = {
      readParagraphs,
      readParagraphPage: vi.fn(async () => {
        throw new Error('Paged host read failed');
      }),
      searchText: vi.fn(async () => []),
    } as unknown as WordHost;
    const { client, contexts } = recordingClient();
    const session = new AssistSession(new WordBridge(host), client, {
      unit: { connectors: [], surfaceContext: { kind: 'word' } },
      context: { docState: false },
    });
    for await (const _event of session.ask('Apa keputusan rapat?')) void _event;
    expect(readParagraphs).not.toHaveBeenCalled();
    expect(textOf(contexts[0]!)).toContain('"source":"unavailable"');
    expect(textOf(contexts[0]!)).toContain('automatic read failed');
    expect(session.context.size).toBe(0);
  });

  it('bounds a whole-mail Markdown table on the actual model context envelope', async () => {
    const body =
      '<pre>| Code | Description |\n| --- | --- |\n' +
      Array.from({ length: 2000 }, (_, i) => `| R${i} | ${'🧪漢字\\"'.repeat(20)} |`).join('\n') +
      '\n</pre>';
    const simulator = installFakeOutlook(outlookSeed({ subject: 'Facilities review', body }));
    try {
      const { client, contexts } = recordingClient();
      const session = new AssistSession(new OutlookBridge(), client, {
        unit: { connectors: [], surfaceContext: { kind: 'outlook' } },
        context: { docState: false },
      });
      for await (const _event of session.ask('Summarise this email')) void _event;
      const parts = contexts[0]!;
      expect(new TextEncoder().encode(JSON.stringify(parts)).length).toBeLessThanOrEqual(24 * 1024);
      const text = textOf(parts);
      expect(text).toContain('| R0 |');
      expect(text).not.toContain('| R1999 |');
      expect(text).toContain('"clippedParts":1');
      expect(text).toContain('"read":"read"');
      expect(text).toContain('not a continuation of this capture');
      expect(text).not.toContain('\ufffd');
      expect(session.context.size).toBe(0);
      expect(simulator.snapshot().replyForms).toHaveLength(0);
    } finally {
      simulator.restore();
    }
  });

  it('retains a targeted mail hit when the supplemental whole-item read fails', async () => {
    const simulator = installFakeOutlook(
      outlookSeed({
        subject: 'Facilities review',
        body: '<p>The Zircon deadline is December 9.</p>',
      }),
    );
    try {
      const bridge = new OutlookBridge();
      vi.spyOn(bridge, 'resolveContext').mockRejectedValue(new Error('Supplemental read failed'));
      const { client, contexts } = recordingClient();
      const session = new AssistSession(bridge, client, {
        unit: { connectors: [], surfaceContext: { kind: 'outlook' } },
        context: { docState: false },
      });
      for await (const _event of session.ask('What is the Zircon deadline?')) void _event;
      const text = textOf(contexts[0]!);
      expect(text).toContain('The Zircon deadline is December 9.');
      expect(text).toContain('"source":"unavailable"');
      expect(session.context.size).toBe(0);
    } finally {
      simulator.restore();
    }
  });
});
