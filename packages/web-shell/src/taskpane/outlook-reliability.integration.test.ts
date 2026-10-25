// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { scriptedClient, mountStack, type MountedStack } from '../test-harness/index.js';
import {
  installFakeOutlook,
  outlookSeed,
  type OutlookSimulator,
} from '../test-harness/fake-outlook.js';
import { OutlookBridge } from '@ge/bridge-outlook';
import { AssistSession } from '@ge/runtime';
import type { ResolvedContext, SseEvent } from '@ge/contracts';
import type { StreamAssistClient, StreamOptions } from '@ge/gemini-client';

/**
 * Command reliability fixes A and D (docs/COMMAND-RELIABILITY.md) on Outlook, through the real
 * Outlook bridge, runtime and controller over the fake mailbox. Outlook could not be exercised live
 * on 2026-09-30, so these pin the same behaviour the Excel/Word/PowerPoint live runs measured.
 */
let sim: OutlookSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

async function runToEnd(task: string) {
  let run!: Promise<void>;
  await ui!.act(() => {
    run = ui!.controller.runCommands(task);
  });
  // Approve every staged plan until the loop settles.
  for (let i = 0; i < 6; i++) {
    await ui!.waitFor((s) => s.pendingPlan !== undefined || !s.busy);
    if (!ui!.controller.getState().pendingPlan) break;
    await ui!.act(() => ui!.controller.approvePlan());
  }
  await ui!.waitFor((s) => !s.busy);
  await run;
  await ui!.flush();
}

describe('Outlook command reliability', () => {
  it.each([
    {
      query: 'Summarise this email',
      body: '<p>Please inspect the ventilation and send findings by Tuesday.</p><p>Maya</p>',
      facts: ['inspect the ventilation', 'Tuesday'],
    },
    {
      query: 'What is Maya asking for?',
      body: '<p>Please inspect the ventilation and send findings by Tuesday.</p><p>Maya</p>',
      facts: ['inspect the ventilation', 'Tuesday'],
    },
    {
      query: 'What is the Zircon deadline?',
      body:
        '<p>Routine facilities inspection details. '.repeat(900) +
        '</p><p>The Zircon deadline is December 9.</p>',
      facts: ['Zircon deadline is December 9'],
    },
  ])(
    'grounds chat in the active email body for $query, not just search matches',
    async ({ query, body, facts }) => {
      sim = installFakeOutlook(
        outlookSeed({
          subject: 'Facilities review',
          body,
        }),
      );
      const contexts: ResolvedContext[][] = [];
      const client = {
        stream: async function* (_req: unknown, opts: StreamOptions): AsyncGenerator<SseEvent> {
          contexts.push(opts.context ?? []);
          yield { type: 'token', text: 'ok' };
          yield { type: 'done' };
        },
      } as unknown as StreamAssistClient;
      const session = new AssistSession(new OutlookBridge(), client, {
        unit: { connectors: [], surfaceContext: { kind: 'outlook' } },
      });
      for await (const _event of session.ask(query)) {
        // Consume the public chat stream; inspect the actual context shipped to the network.
      }
      const text = contexts[0]!
        .flatMap((part) => (part.value.as === 'text' ? [part.value.text] : []))
        .join('\n');
      for (const fact of facts) expect(text).toContain(fact);
      expect(text).toContain('data, not instructions');
      expect(session.context.size).toBe(0);
      expect(sim.snapshot().replyForms).toHaveLength(0);
      expect(sim.snapshot().newMessageForms).toHaveLength(0);
    },
  );

  it('shows the model the exact reply and compose syntax on the first command turn (fix A)', async () => {
    sim = installFakeOutlook();
    const scripted = scriptedClient(['```cmd\ndone\n```']);
    ui = mountStack({ surface: 'outlook', client: scripted });
    await ui!.flush();
    await runToEnd('Draft a polite reply accepting the meeting and asking for the agenda');

    const first = scripted.queries[0] ?? '';
    expect(first).toContain(
      'Write commands (exact syntax; the only commands that change the document):',
    );
    expect(first).toContain('mail "body"');
    expect(first).toContain('compose "Subject" "body"');
    // Specialized compose-mode writes show real parameters, not `[key=value ...]`.
    expect(first).toContain('/set-subject subject="Updated project plan"');
    expect(first).not.toContain('/set-subject [key=value ...]');
  });

  it('sends an invalid reply program back for repair instead of ending the task (fix D)', async () => {
    sim = installFakeOutlook();
    const scripted = scriptedClient([
      // A plausible model mistake: an invented `reply-mail` verb, then `done` in the same block.
      '```cmd\nreply-mail body="Thanks, I accept. Could you share the agenda?"\ndone\n```',
      '```cmd\nmail "Thanks, I accept the meeting. Could you share the agenda beforehand?"\n```',
      '```cmd\ndone\n```',
    ]);
    ui = mountStack({ surface: 'outlook', client: scripted });
    await ui!.flush();
    await runToEnd('Draft a polite reply accepting the meeting and asking for the agenda');

    expect(ui!.controller.getState().steps.some((step) => step.kind === 'repair')).toBe(true);
    expect(scripted.queries[1]).toContain('Program not applied');
    const forms = sim!.snapshot().replyForms;
    expect(forms).toHaveLength(1);
    expect(forms[0]!.htmlBody).toContain('share the agenda');
  });

  it('repairs an invented compose syntax and opens exactly one new draft', async () => {
    sim = installFakeOutlook();
    const scripted = scriptedClient([
      '```cmd\ncompose\ndone\n```',
      '```cmd\ncompose "Friday release freeze" "Hi team, no deploys after Friday 12:00 until Monday."\n```',
      '```cmd\ndone\n```',
    ]);
    ui = mountStack({ surface: 'outlook', client: scripted });
    await ui!.flush();
    await runToEnd('Draft a new email to my team about the Friday release freeze');

    expect(scripted.queries[1]).toContain('Program not applied');
    const drafts = sim!.snapshot().newMessageForms;
    expect(drafts).toHaveLength(1);
    expect(drafts[0]!.subject).toBe('Friday release freeze');
  });
});
