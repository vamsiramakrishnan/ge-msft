import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import { asChangeId } from '@ge/contracts';
import { TriggerRegistry, type HostEvent, type Scheduler } from '@ge/triggers';
import { StreamAssistClient } from '@ge/gemini-client';
import { Orchestrator } from './orchestrator.js';
import { AssistSession } from './assist-session.js';
import type { DocBridge } from './bridge.js';

const tick = (): Promise<void> => new Promise((r) => setTimeout(r, 0));

class FakeBridge implements DocBridge {
  readonly surface = 'word' as const;
  applied: ActuationRequest[] = [];
  emit: ((e: HostEvent) => void) | null = null;

  getCapabilities(): CapabilityManifest {
    return {
      surface: 'word',
      contextKinds: [],
      actuations: [{ kind: 'tracked-change', surface: 'word', title: 'Edit', reversible: true }],
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
  watch(emit: (e: HostEvent) => void): () => void {
    this.emit = emit;
    return () => {
      this.emit = null;
    };
  }
}

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'u@acme' };
const dummyClient = new StreamAssistClient(
  { getAccessToken: () => Promise.resolve('t') },
  cfg,
  (async () => new Response('[]', { status: 200 })) as never,
);

describe('Orchestrator — react to events', () => {
  it('routes a non-frequency event immediately to onAutomate', async () => {
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry();
    reg.register({
      id: 'triage',
      on: 'mail-received',
      handle: () => ({ kind: 'automate', query: 'triage this' }),
    });
    const onAutomate = vi.fn();
    const orch = new Orchestrator(bridge, reg, { onAutomate });
    orch.start();

    bridge.emit?.({ type: 'mail-received', id: 'm1' });
    await tick();
    expect(onAutomate).toHaveBeenCalledWith('triage this');
  });

  it('debounces high-frequency selection events into one suggestion', async () => {
    const holder: { fn: (() => void) | null } = { fn: null };
    const scheduler: Scheduler = {
      set: (fn) => ((holder.fn = fn), 1),
      clear: () => (holder.fn = null),
    };
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry();
    reg.register({
      id: 'ambient',
      on: 'selection-changed',
      handle: () => ({ kind: 'suggest', title: 'insight' }),
    });
    const onSuggest = vi.fn();
    const orch = new Orchestrator(bridge, reg, { onSuggest }, { scheduler });
    orch.start();

    bridge.emit?.({ type: 'selection-changed', surface: 'word', origin: 'local', preview: 'a' });
    bridge.emit?.({ type: 'selection-changed', surface: 'word', origin: 'local', preview: 'ab' });
    expect(onSuggest).not.toHaveBeenCalled(); // debounced
    holder.fn?.(); // fire the debounce timer
    await tick();
    expect(onSuggest).toHaveBeenCalledTimes(1);
    expect(onSuggest).toHaveBeenCalledWith({ title: 'insight' });

    orch.stop();
    expect(bridge.emit).toBeNull();
  });

  it('start() is idempotent — a second start does not re-subscribe', () => {
    let watchCount = 0;
    class CountingBridge extends FakeBridge {
      override watch(emit: (e: HostEvent) => void): () => void {
        watchCount += 1;
        return super.watch(emit);
      }
    }
    const bridge = new CountingBridge();
    const orch = new Orchestrator(bridge, new TriggerRegistry(), {});
    orch.start();
    orch.start(); // no-op: already started
    expect(watchCount).toBe(1);
  });

  it('start() is a no-op when the bridge cannot observe events (no watch)', () => {
    // A bridge WITHOUT watch — start must not throw and must not subscribe.
    class NoWatchBridge extends FakeBridge {}
    const bridge = new NoWatchBridge();
    // Remove the inherited watch so the orchestrator's `!this.bridge.watch` guard trips.
    (bridge as { watch?: unknown }).watch = undefined;
    const orch = new Orchestrator(bridge, new TriggerRegistry(), {});
    expect(() => orch.start()).not.toThrow();
    // No watcher was wired, so emit was never assigned.
    expect(bridge.emit).toBeNull();
    // stop() on a never-started orchestrator is safe too.
    expect(() => orch.stop()).not.toThrow();
  });

  it('forwards a suggestion outcome WITH its detail and query through onSuggest', async () => {
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry();
    reg.register({
      id: 'rich',
      on: 'mail-received',
      handle: () => ({
        kind: 'suggest',
        title: 'Triage',
        detail: 'Looks urgent',
        query: 'summarize this thread',
      }),
    });
    const onSuggest = vi.fn();
    const orch = new Orchestrator(bridge, reg, { onSuggest });
    orch.start();

    bridge.emit?.({ type: 'mail-received', id: 'm9' });
    await tick();
    expect(onSuggest).toHaveBeenCalledWith({
      title: 'Triage',
      detail: 'Looks urgent',
      query: 'summarize this thread',
    });
  });

  it('routes every event to onContext, independent of whether a trigger matches', async () => {
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry(); // no triggers registered
    const onContext = vi.fn();
    const orch = new Orchestrator(bridge, reg, { onContext });
    orch.start();

    bridge.emit?.({ type: 'mail-received', id: 'm1' });
    await tick();
    expect(onContext).toHaveBeenCalledWith(
      { type: 'mail-received', id: 'm1' },
      { signal: expect.any(AbortSignal) },
    );
  });
});

describe('AssistSession actuation gate (PreToolUse analog)', () => {
  const unit = { connectors: [], surfaceContext: { kind: 'word' as const } };

  it('blocks a write when a pre-actuation trigger vetoes', async () => {
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry();
    reg.register({
      id: 'guard',
      on: 'pre-actuation',
      handle: () => ({ kind: 'block', reason: 'Not grounded.' }),
    });
    const session = new AssistSession(bridge, dummyClient, { unit, triggers: reg });

    const result = await session.apply('tracked-change', { text: 'x' }, asChangeId('c1'));
    expect(result).toMatchObject({
      ok: false,
      error: { code: 'blocked', message: 'Not grounded.' },
    });
    expect(bridge.applied).toHaveLength(0); // write never reached the host
  });

  it('applies and fires post-actuation audit when no gate blocks', async () => {
    const bridge = new FakeBridge();
    const reg = new TriggerRegistry();
    const audit = vi.fn(() => ({ kind: 'continue' }) as const);
    reg.register({ id: 'audit', on: 'post-actuation', handle: audit });
    const session = new AssistSession(bridge, dummyClient, { unit, triggers: reg });

    const result = await session.apply('tracked-change', { text: 'ok' }, asChangeId('c2'));
    expect(result.ok).toBe(true);
    expect(bridge.applied).toHaveLength(1);
    await tick();
    expect(audit).toHaveBeenCalledOnce();
  });
});

describe('Orchestrator event ownership', () => {
  it('keeps document and selection debounce independent and cancels queued events on stop', async () => {
    const timers = new Map<number, () => void>();
    let next = 0;
    const scheduler: Scheduler = {
      set(fn) {
        const id = ++next;
        timers.set(id, fn);
        return id;
      },
      clear(id) {
        timers.delete(id as number);
      },
    };
    const bridge = new FakeBridge();
    const seen: HostEvent[] = [];
    const orch = new Orchestrator(
      bridge,
      new TriggerRegistry(),
      {
        onContext(event) {
          seen.push(event);
        },
      },
      { scheduler },
    );
    orch.start();
    bridge.emit?.({ type: 'selection-changed', surface: 'word', origin: 'local', preview: 'a' });
    bridge.emit?.({ type: 'document-changed', surface: 'word', origin: 'local' });
    bridge.emit?.({ type: 'selection-changed', surface: 'word', origin: 'local', preview: 'b' });
    expect(timers.size).toBe(2);
    for (const timer of timers.values()) timer();
    timers.clear();
    await orch.idle();
    expect(seen.map((e) => e.type)).toEqual(['document-changed', 'selection-changed']);
    expect(seen[1]).toMatchObject({ preview: 'b' });
    bridge.emit?.({ type: 'document-changed', surface: 'word', origin: 'local' });
    const stale = [...timers.values()];
    orch.stop();
    expect(timers.size).toBe(0);
    orch.start();
    for (const timer of stale) timer();
    await orch.idle();
    expect(seen).toHaveLength(2);
    orch.stop();
  });

  it('aborts context work on stop and suppresses its downstream reactions', async () => {
    const bridge = new FakeBridge();
    const registry = new TriggerRegistry();
    registry.register({
      id: 'suggest',
      on: 'mail-received',
      handle: () => ({ kind: 'suggest', title: 'Old message' }),
    });
    let contextSignal: AbortSignal | undefined;
    let entered!: () => void;
    const started = new Promise<void>((r) => {
      entered = r;
    });
    const onSuggest = vi.fn();
    const orch = new Orchestrator(bridge, registry, {
      onContext(_event, { signal }) {
        contextSignal = signal;
        entered();
        return new Promise(() => {});
      },
      onSuggest,
    });
    orch.start();
    orch.publish({ type: 'mail-received', id: 'old' });
    await started;
    orch.stop();
    await orch.idle();
    expect(contextSignal?.aborted).toBe(true);
    expect(onSuggest).not.toHaveBeenCalled();
  });
});
