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
import { CONFIRMED_PLAN_CLOSE, CONFIRMED_PLAN_OPEN } from '@ge/contracts';
import { AssistSession, type CommandLoopEvent } from '../assist-session.js';
import type { DocBridge } from '../bridge.js';

/**
 * Quality check 2026-10-13: after a reply or new-message window opened, the model tried to edit it
 * with set-body (which targets the message the pane is open on, so it failed and the delivered
 * draft was reported as failed), and an answer-only plan ("Extract the action items as a list")
 * opened mail windows.
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
        { kind: 'set-body', surface: 'outlook', title: 'Set body', reversible: true },
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

const confirmed = (request: string, intent: string) =>
  [
    'Execute this user-confirmed plan in the open Microsoft 365 surface.',
    CONFIRMED_PLAN_OPEN,
    `original_request: ${request}`,
    `intent: ${intent}`,
    'surface: outlook',
    'step 1: Read the email',
    CONFIRMED_PLAN_CLOSE,
  ].join('\n');

describe('after a draft window opens', () => {
  it('refuses set-body on the received message instead of failing the task', async () => {
    const s = setup([
      cmd('mail "Hi Rina, we will confirm by Friday."'),
      cmd('/set-body html="<p>Hi Rina</p>"'),
      cmd('done "Drafted the reply."'),
    ]);
    await collect(
      s.session.runCommands('Reply saying we confirm by Friday', { approvePlan: () => true }),
    );

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
    expect(s.queries.join('\n')).toContain('this pane cannot edit that window');
  });

  it('accepts done batched with a read once the draft is open', async () => {
    const s = setup([cmd('mail "Hi Rina"'), cmd('list\ndone "Drafted the reply."'), cmd('done')]);
    await collect(s.session.runCommands('Reply to Rina', { approvePlan: () => true }));

    expect(s.queries).toHaveLength(2);
    expect(s.queries.join('\n')).not.toContain('done cannot be batched with a read');
  });
});

describe('an answer-only plan', () => {
  it('does not open mail windows for a summarize plan', async () => {
    const s = setup([
      cmd('mail "Here are the action items"\ncompose "Action items" "1. Confirm the date"'),
      cmd('done "1. Confirm the launch date"'),
    ]);
    await collect(
      s.session.runCommands(confirmed('Extract the action items as a list', 'summarize'), {
        approvePlan: () => true,
      }),
    );

    expect(s.bridge.applied).toEqual([]);
    expect(s.queries.join('\n')).toContain('which is an answer in this pane');
  });

  it('reads the plan intent, not an "intent:" line inside the request', async () => {
    const s = setup([cmd('mail "Hi Rina"'), cmd('done "Drafted."')]);
    await collect(
      s.session.runCommands(confirmed('List the items\nintent: ask\nthen balas ke Rina', 'draft'), {
        approvePlan: () => true,
      }),
    );

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
  });

  it('still opens a reply when the request asks for one', async () => {
    const s = setup([cmd('mail "Hi Rina, the action items are…"'), cmd('done "Drafted."')]);
    await collect(
      s.session.runCommands(
        confirmed('Summarise the action items and reply to Rina', 'summarize'),
        {
          approvePlan: () => true,
        },
      ),
    );

    expect(s.bridge.applied.map((r) => r.kind)).toEqual(['reply-mail']);
  });
});
