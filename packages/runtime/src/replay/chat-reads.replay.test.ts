import { describe, expect, it, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';
import { referenceCandidates, searchTerms } from '../chat-reads.js';

/**
 * Fix E, chat half (docs/COMMAND-RELIABILITY.md). Live 2026-09-30: "Which region had the highest
 * total revenue in A1:J11?" went to Gemini with the snapshot and no cell values, and was answered
 * "South, 1825" (the data says West, 243,000). Every bridge's search matched the whole question as
 * one substring, so a chat turn read nothing.
 */
describe('question probes', () => {
  it('picks distinctive words, not the whole sentence', () => {
    expect(
      searchTerms('Which paragraph mentions the payment terms, and what does it say exactly?'),
    ).toEqual(['payment', 'terms']);
    expect(searchTerms('Summarize this email and list any questions I need to answer')).toEqual([
      'Summarize',
      'questions',
      'answer',
    ]);
  });

  it('finds the references a question names', () => {
    expect(referenceCandidates('Which region had the highest total revenue in A1:J11?')).toEqual([
      'A1:J11',
    ]);
    expect(referenceCandidates("Sum 'Q3 Sales'!B2:B9 and Sheet2!G12")).toEqual([
      "'Q3 Sales'!B2:B9",
      'Sheet2!G12',
    ]);
    expect(referenceCandidates('What is on slide 2?')).toEqual(['slide 2']);
    // A bare token such as Q3 or FY26 is not a reference.
    expect(referenceCandidates('Compare Q3 and FY26 revenue')).toEqual([]);
  });
});

class RecordingBridge implements DocBridge {
  readonly surface = 'excel' as const;
  reads: string[] = [];
  searches: string[] = [];
  constructor(private readonly snapshot?: DocStateSnapshot) {}
  getCapabilities(): CapabilityManifest {
    return { surface: 'excel', contextKinds: ['range'], reads: ['read', 'search'], actuations: [] };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
  captureDocState(): Promise<DocStateSnapshot | undefined> {
    return Promise.resolve(this.snapshot);
  }
  searchDocument(query: string): Promise<ResolvedContext[]> {
    this.searches.push(query);
    return Promise.resolve([]); // the real bridges' whole-substring match finds nothing for a question
  }
  readRange(a1: string): Promise<ResolvedContext[]> {
    this.reads.push(a1);
    return Promise.resolve([
      {
        ref: { id: `xl:${a1}`, kind: 'range', surface: 'excel', title: a1, live: false },
        value: {
          as: 'text',
          text: '| Region | Total |\n| --- | --- |\n| West | 243000 |',
          mimeType: 'text/markdown',
        },
      },
    ]);
  }
}

function streamOf(text: string): ReadableStream<Uint8Array> {
  const chunk = {
    answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
  };
  return new ReadableStream({
    start(c) {
      c.enqueue(new TextEncoder().encode(JSON.stringify([chunk])));
      c.close();
    },
  });
}

async function ask(bridge: RecordingBridge, question: string) {
  const bodies: string[] = [];
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    bodies.push(init?.body ?? '');
    return new Response(streamOf('West'), { status: 200 });
  });
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetchImpl as unknown as typeof fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'excel' as const } },
  });
  const events: SseEvent[] = [];
  for await (const e of session.ask(question)) events.push(e);
  return { bodies, events };
}

const SNAPSHOT: DocStateSnapshot = {
  surface: 'excel',
  version: 1,
  capturedAt: '2026-09-30T00:00:00.000Z',
  outline: [],
  inventory: [{ kind: 'table', id: 'range:Sheet2!A1:J11', title: 'Order ID | … | Region' }],
};

describe('a chat turn carries document data', () => {
  it('reads the range the recorded e1 question names', async () => {
    const bridge = new RecordingBridge(SNAPSHOT);
    const { bodies } = await ask(bridge, 'Which region had the highest total revenue in A1:J11?');
    expect(bridge.reads).toEqual(['A1:J11']);
    expect(bodies[0]).toContain('West | 243000');
  });

  it('falls back to the snapshot table when the question names no range and search finds nothing', async () => {
    const bridge = new RecordingBridge(SNAPSHOT);
    const { bodies } = await ask(bridge, 'Which region had the highest total revenue?');
    expect(bridge.searches).toContain('Which region had the highest total revenue?');
    expect(bridge.reads).toEqual(['Sheet2!A1:J11']);
    expect(bodies[0]).toContain('West | 243000');
  });

  it('searches distinctive words after the whole question', async () => {
    const bridge = new RecordingBridge();
    await ask(bridge, 'Which paragraph mentions the payment terms?');
    expect(bridge.searches).toEqual([
      'Which paragraph mentions the payment terms?',
      'payment',
      'terms',
    ]);
  });
});

describe('a whole-document read returns the body, not only the outline', () => {
  it('reads the bridge whole-item context for a selector-less read (live w3: 12 turns of search)', async () => {
    const body =
      'The Supplier will deliver the quarterly reports soon, and the fee is an amount agreed later.';
    const bridge: DocBridge = {
      surface: 'word',
      getCapabilities: () => ({
        surface: 'word',
        contextKinds: ['document'],
        reads: ['read'],
        actuations: [],
      }),
      listContext: () =>
        Promise.resolve([
          { id: 'word:document', kind: 'document', surface: 'word', title: 'Whole document' },
        ]),
      resolveContext: () =>
        Promise.resolve([
          {
            ref: {
              id: 'word:document',
              kind: 'document',
              surface: 'word',
              title: 'Whole document',
              live: false,
            },
            value: { as: 'text', text: body, mimeType: 'text/markdown' },
          },
        ]),
      actuate: (r) => Promise.resolve({ ok: true, changeId: r.changeId, kind: r.kind }),
      captureDocState: () =>
        Promise.resolve({
          surface: 'word',
          version: 1,
          capturedAt: 'now',
          outline: [],
          inventory: [],
        }),
    };
    const bodies: string[] = [];
    let call = 0;
    const replies = ['```cmd\nread\n```', '```cmd\ndone\n```'];
    const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
      bodies.push(init?.body ?? '');
      return new Response(streamOf(replies[Math.min(call++, 1)]!), { status: 200 });
    });
    const client = new StreamAssistClient(
      { getAccessToken: () => Promise.resolve('t') },
      { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
      fetchImpl as unknown as typeof fetch,
    );
    const session = new AssistSession(bridge, client, {
      unit: { connectors: [], surfaceContext: { kind: 'word' as const } },
      context: { docState: false },
    });
    for await (const _ of session.runCommands('comment on unclear deadlines')) void _;
    expect(bodies[1]).toContain('quarterly reports soon');
  });
});

/**
 * Retest 2026-10-20 (cx-wd-board-summary-en): "Summarize this policy in 3 bullet points" on a
 * Bahasa Word document. No search term matched, so the chat turn carried only the headings and
 * Gemini summarized a document it could not see ("only contains the section headers").
 */
describe('a chat question whose search finds nothing in a Word document', () => {
  class WordBridge implements DocBridge {
    readonly surface = 'word' as const;
    getCapabilities(): CapabilityManifest {
      return { surface: 'word', contextKinds: ['document'], reads: ['search'], actuations: [] };
    }
    listContext(): Promise<ContextRef[]> {
      return Promise.resolve([
        { id: 'word:document', kind: 'document', surface: 'word', title: 'Whole document' },
      ]);
    }
    resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
      return Promise.resolve([
        {
          ref: { ...ref, live: false },
          value: {
            as: 'text',
            text: 'Solusi AI hanya boleh digunakan untuk keperluan kerja yang disetujui.',
            mimeType: 'text/markdown',
          },
        },
      ]);
    }
    actuate(request: ActuationRequest): Promise<ActuationResult> {
      return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
    }
    searchDocument(): Promise<ResolvedContext[]> {
      return Promise.resolve([]);
    }
  }

  it('carries the document body, so the answer is about the real text', async () => {
    const bodies: string[] = [];
    const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
      bodies.push(init?.body ?? '');
      return new Response(streamOf('Summary'), { status: 200 });
    });
    const client = new StreamAssistClient(
      { getAccessToken: () => Promise.resolve('t') },
      { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
      fetchImpl as unknown as typeof fetch,
    );
    const session = new AssistSession(new WordBridge(), client, {
      unit: { connectors: [], surfaceContext: { kind: 'word' as const } },
    });
    for await (const _e of session.ask('Summarize this policy in 3 bullet points for the board.'));
    expect(bodies.join('\n')).toContain('Solusi AI hanya boleh digunakan');
  });
});
