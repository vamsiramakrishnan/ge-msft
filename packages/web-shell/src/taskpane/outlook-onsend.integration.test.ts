// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { TriggerRegistry, type HostEvent, type TriggerOutcome } from '@ge/triggers';
import {
  createMessageSendHandler,
  activeItemIdResolver,
  type OnSendCompletedOptions,
} from '@ge/bridge-outlook';
import { scriptedClient, mountStack, type MountedStack } from '../test-harness/index.js';
import {
  installFakeOutlook,
  outlookSeed,
  type OutlookSimulator,
} from '../test-harness/fake-outlook.js';

/**
 * FLAGSHIP full-stack Outlook on-send interplay. Wires the REAL packages —
 * `selectBridge('outlook')` → real {@link "@ge/runtime"!AssistSession} (over the scripted client) →
 * real `PanelController` → rendered `<App/>` for the capture + reviewable-reply path, AND the REAL
 * Outlook on-send glue (`createMessageSendHandler`) over a REAL `TriggerRegistry` for the gate — all
 * against an in-memory fake Outlook host (`installFakeOutlook`, which fakes only the slice of
 * `Office.context.mailbox` the bridge drives). The ONLY mocks are the host object model and the
 * model stream.
 *
 * This is the most consequential interplay in the product: a draft must never leave the client
 * without the user's approval, the gate must release once approved, and — crucially — a crash while
 * SIGNALLING a decided block must NOT silently downgrade to an allow (no silent send). Each test
 * asserts OBSERVABLE cross-boundary behavior: a reply form actually opened on the fake host, the
 * `event.completed` call the host received (`allowEvent`), the body content the gate let through.
 */

let sim: OutlookSimulator | undefined;
let ui: MountedStack | undefined;

afterEach(() => {
  ui?.unmount();
  sim?.restore();
  ui = undefined;
  sim = undefined;
});

/** Collect the `event.completed(...)` calls a handler makes — the host-visible decision. */
function captureSend(): {
  calls: OnSendCompletedOptions[];
  event: { completed: (o?: OnSendCompletedOptions) => void };
} {
  const calls: OnSendCompletedOptions[] = [];
  return { calls, event: { completed: (o?: OnSendCompletedOptions) => calls.push(o ?? {}) } };
}

describe('Outlook on-send full-stack interplay', () => {
  it('captures the mail item and opens a reviewable reply form ONLY after plan approval', async () => {
    sim = installFakeOutlook();
    ui = mountStack({
      surface: 'outlook',
      client: scriptedClient([
        '```cmd\nmail "Thanks — we can hold the SLA at 99.5%; the liability cap stays at 6 months. Happy to discuss Friday."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('draft a grounded reply to the vendor');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);

    // Nothing has been written to the host while the plan is still pending approval.
    expect(sim!.snapshot().replyForms.length).toBe(0);

    // The plan-approval card stages the reply effect for review (the reply-mail actuation kind).
    // (The reply BODY lands on the host below — that is the load-bearing cross-boundary assertion.)
    const planCmd = ui!.container.querySelector('.plan-approval .cmd')?.textContent ?? '';
    expect(planCmd).toContain('reply-mail');

    await ui!.act(() => ui!.controller.approvePlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    // The REAL bridge opened a reviewable reply form on the fake host (a reply, never a send).
    const snap = sim!.snapshot();
    expect(snap.replyForms.length).toBe(1);
    expect(snap.replyForms[0]?.htmlBody).toContain('liability cap stays at 6 months');
    expect(snap.newMessageForms.length).toBe(0);
  });

  it('rejecting the plan opens NO reply form (the reviewable draft never leaves the panel)', async () => {
    sim = installFakeOutlook();
    ui = mountStack({
      surface: 'outlook',
      client: scriptedClient([
        '```cmd\nmail "Approving the 12-month liability cap as requested."\n```',
        '```cmd\ndone\n```',
      ]),
    });
    await ui!.flush();

    let run!: Promise<void>;
    await ui!.act(() => {
      run = ui!.controller.runCommands('agree to their terms');
    });
    await ui!.waitFor((s) => s.pendingPlan !== undefined);
    await ui!.act(() => ui!.controller.rejectPlan());
    await ui!.waitFor((s) => !s.busy);
    await run;
    await ui!.flush();

    expect(sim!.snapshot().replyForms.length).toBe(0);
  });

  it('the on-send gate BLOCKS the send while approval is pending, then ALLOWS it after approval', async () => {
    // The fake host gives the resolver a saved itemId — the gate sees the real active draft id.
    sim = installFakeOutlook(
      outlookSeed({
        itemId: 'AAMk-draft-42',
        subject: 'Re: liability cap',
        from: { emailAddress: 'procurement@northwind.example' },
        body: '<p>Please confirm.</p>',
      }),
    );

    // A REAL trigger registry with a stateful guard: an outbound reply needs explicit approval. The
    // gate reads the real `mail-send` event the on-send glue builds (carrying the resolved itemId).
    const registry = new TriggerRegistry();
    let approved = false;
    let gatedId: string | undefined = 'unset';
    registry.register({
      id: 'reply-needs-approval',
      on: 'mail-send',
      handle: (e: HostEvent): TriggerOutcome => {
        gatedId = e.type === 'mail-send' ? e.id : undefined;
        return approved
          ? { kind: 'continue' }
          : { kind: 'block', reason: 'This grounded reply needs review before it is sent.' };
      },
    });

    const handler = createMessageSendHandler(registry, { resolveItemId: activeItemIdResolver });

    // 1) Pending approval → the gate BLOCKS: the host is told NOT to send, with the reason surfaced.
    const blocked = captureSend();
    await handler(blocked.event);
    expect(gatedId).toBe('AAMk-draft-42'); // the resolver read the real active item id end to end
    expect(blocked.calls).toEqual([
      { allowEvent: false, errorMessage: 'This grounded reply needs review before it is sent.' },
    ]);

    // 2) User approves → the same gate now ALLOWS the send.
    approved = true;
    const allowed = captureSend();
    await handler(allowed.event);
    expect(allowed.calls).toEqual([{ allowEvent: true }]);
  });

  it('a registered required check failure blocks Send with a recoverable reason', async () => {
    sim = installFakeOutlook();
    const registry = new TriggerRegistry();
    // A registered required check must finish before Send can proceed.
    registry.register({
      id: 'buggy-guard',
      on: 'mail-send',
      handle: (): TriggerOutcome => {
        throw new Error('guard exploded while deciding');
      },
    });
    const handler = createMessageSendHandler(registry);

    const send = captureSend();
    await handler(send.event);
    // Exactly one completion; do not expose raw exception details or silently bypass the check.
    expect(send.calls).toEqual([
      {
        allowEvent: false,
        errorMessage:
          'Required check buggy-guard could not complete. Try again before applying this action.',
      },
    ]);
  });

  it('a crash while SIGNALLING a decided BLOCK does NOT silently downgrade to an allow (no silent send)', async () => {
    sim = installFakeOutlook();
    const registry = new TriggerRegistry();
    registry.register({
      id: 'always-block',
      on: 'mail-send',
      handle: (): TriggerOutcome => ({ kind: 'block', reason: 'blocked pending review' }),
    });
    const handler = createMessageSendHandler(registry);

    // The host's `completed` throws on the FIRST call (the block signal). The fail-safe must NOT
    // re-call completed({ allowEvent: true }) — that would silently send a blocked mail. Instead the
    // error propagates: the decision was already a block, so it must not be downgraded.
    let completedCalls = 0;
    const seen: OnSendCompletedOptions[] = [];
    const event = {
      completed: (o?: OnSendCompletedOptions): void => {
        completedCalls += 1;
        seen.push(o ?? {});
        throw new Error('host completed() failed mid-block');
      },
    };

    await expect(handler(event)).rejects.toThrow(/host completed\(\) failed mid-block/);
    // The handler attempted exactly the BLOCK signal and did NOT retry with an allow.
    expect(completedCalls).toBe(1);
    expect(seen).toEqual([{ allowEvent: false, errorMessage: 'blocked pending review' }]);
  });

  it('the bridge watch() emits a mail-send-eligible draft event when the active item changes', async () => {
    // Tie the capture path to the gate path: switching to a saved item makes its id resolvable, so a
    // subsequent on-send gate sees that same id. Drives the REAL bridge.watch() over the fake host.
    sim = installFakeOutlook(
      outlookSeed({
        itemId: 'AAMk-switched-7',
        subject: 'Re: renewal',
        body: '<p>body</p>',
      }),
    );
    ui = mountStack({ surface: 'outlook', client: scriptedClient([]) });
    await ui!.flush();

    const events: HostEvent[] = [];
    const watch = ui!.bridge.watch;
    expect(watch).toBeDefined();
    const unsubscribe = watch!.call(ui!.bridge, (e: HostEvent) => events.push(e));
    // The host fires ItemChanged — the user switched to the saved item.
    sim!.mailbox.fire(sim!.eventType.ItemChanged);

    // The bridge classified the saved item as a received (read-mode) event carrying its opaque id.
    expect(events.length).toBe(1);
    const ev = events[0]!;
    expect(ev.type).toBe('mail-received');
    expect(ev.type === 'mail-received' ? ev.id : undefined).toBe('AAMk-switched-7');

    // And the same id is what the on-send resolver reads for the gate — capture ↔ gate share it.
    expect(activeItemIdResolver()).toBe('AAMk-switched-7');

    unsubscribe();
    expect(sim!.mailbox.count(sim!.eventType.ItemChanged)).toBe(0);
  });
});
