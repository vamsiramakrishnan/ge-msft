import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
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
    return { surface: 'word', contextKinds: [], actuations: [] };
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

    const result = await session.apply('tracked-change', { text: 'x' }, 'c1');
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

    const result = await session.apply('tracked-change', { text: 'ok' }, 'c2');
    expect(result.ok).toBe(true);
    expect(bridge.applied).toHaveLength(1);
    await tick();
    expect(audit).toHaveBeenCalledOnce();
  });
});
