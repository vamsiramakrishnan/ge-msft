import { describe, expect, it, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
  SseEvent,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';

/**
 * Live scenarios 2026-10-09: "Write an email to the vendor asking for an update" opened four reply
 * windows. A reply or new-message window cannot be read back, so the model kept opening another.
 */
const cmd = (program: string) => '```cmd\n' + program + '\n```';

class OutlookLikeBridge implements DocBridge {
  readonly surface = 'outlook' as const;
  applied: ActuationRequest[] = [];
  getCapabilities(): CapabilityManifest {
    return {
      surface: 'outlook',
      contextKinds: ['mail-item', 'document'],
      reads: ['read'],
      actuations: [
        { kind: 'reply-mail', surface: 'outlook', title: 'Reply', reversible: false },
        { kind: 'create-mail', surface: 'outlook', title: 'New email', reversible: false },
      ],
    };
  }
  listContext(): Promise<ContextRef[]> {
    return Promise.resolve([]);
  }
  resolveContext(): Promise<ResolvedContext[]> {
    return Promise.resolve([]);
  }
  actuate(request: ActuationRequest): Promise<ActuationResult> {
    this.applied.push(request);
    return Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind });
  }
}

function setup(turns: string[]) {
  const bridge = new OutlookLikeBridge();
  const queries: string[] = [];
  let call = 0;
  const fetchImpl = vi.fn(async (_url: string, init?: { body?: string }) => {
    const body = JSON.parse(init?.body ?? '{}') as { query?: { text?: string } };
    queries.push(body.query?.text ?? '');
    const text = turns[Math.min(call++, turns.length - 1)] ?? cmd('done');
    const chunk = {
      sessionInfo: { session: 's' },
      answer: { state: 'SUCCEEDED', replies: [{ groundedContent: { content: { text } } }] },
    };
    const enc = new TextEncoder().encode(JSON.stringify([chunk]));
    return new Response(
      new ReadableStream({
        start(c) {
          c.enqueue(enc);
          c.close();
        },
      }),
      { status: 200 },
    );
  });
  const client = new StreamAssistClient(
    { getAccessToken: () => Promise.resolve('t') },
    { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' },
    fetchImpl as unknown as typeof fetch,
  );
  const session = new AssistSession(bridge, client, {
    unit: { connectors: [], surfaceContext: { kind: 'outlook' as const } },
    context: { docState: false },
  });
  return { bridge, queries, session };
}

async function collect(gen: AsyncGenerator<SseEvent | CommandLoopEvent>): Promise<void> {
  for await (const _e of gen) void _e;
}

describe('one reply or new-message window per task', () => {
  it('refuses a second reply in the same task and tells the model to finish', async () => {
    const s = setup([
      cmd('mail "Hi Rina, any update?"'),
      cmd('mail "Hi Rina, any update please?"'),
      cmd('done'),
    ]);
    await collect(s.session.runCommands('Reply asking for an update', { approvePlan: () => true }));

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
    expect(s.queries.join('\n')).toContain('this second one was not opened');
  });

  it('still allows one reply and one new email in the same task', async () => {
    const s = setup([cmd('mail "Thanks Rina"\ncompose "Update" "Hi Budi"'), cmd('done')]);
    await collect(
      s.session.runCommands('Reply to Rina and email Budi', { approvePlan: () => true }),
    );

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail', 'create-mail']);
  });

  it('refuses a second reply in the same command program', async () => {
    const s = setup([cmd('mail "Hi Rina"\nmail "Hi again Rina"'), cmd('done')]);
    await collect(s.session.runCommands('Reply to Rina', { approvePlan: () => true }));

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
  });
});
