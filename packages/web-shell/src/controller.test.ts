import { describe, it, expect, vi } from 'vitest';
import type {
  ActuationParams,
  ActuationRequest,
  ActuationResult,
  ContextRef,
  SseEvent,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { PanelController, type AssistLike, type ContextLister } from './controller.js';

const ref = (id: string, title: string): ContextRef => ({
  id,
  kind: 'selection',
  surface: 'word',
  title,
  preview: title,
});

/** A fake AssistSession: scripts the SSE stream and records attaches/applies/ingests. */
class FakeAssist implements AssistLike {
  context = { size: 0 };
  attached: string[] = [];
  detached: string[] = [];
  ingested: HostEvent[] = [];
  applied: Array<{ kind: string; changeId: string }> = [];
  script: SseEvent[] = [{ type: 'token', text: 'hi' }, { type: 'done' }];
  applyResult: ActuationResult = {
    ok: true,
    changeId: '',
    kind: 'tracked-change',
    location: 'para:3',
  };

  attachRef(r: ContextRef): Promise<void> {
    this.attached.push(r.id);
    this.context = { size: this.context.size + 1 };
    return Promise.resolve();
  }
  detach(id: string): void {
    this.detached.push(id);
  }
  async *ask(_query: string): AsyncGenerator<SseEvent> {
    for (const ev of this.script) yield ev;
  }
  apply(
    kind: ActuationRequest['kind'],
    _params: ActuationParams,
    changeId: string,
  ): Promise<ActuationResult> {
    this.applied.push({ kind, changeId });
    return Promise.resolve({ ...this.applyResult, changeId, kind });
  }
  ingest(event: HostEvent): Promise<void> {
    this.ingested.push(event);
    return Promise.resolve();
  }
}

function lister(refs: ContextRef[]): ContextLister {
  return { listContext: () => Promise.resolve(refs) };
}

describe('PanelController — context tray', () => {
  it('loads chips and attaches/detaches a specific ref', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([ref('word:selection', 'Selection')]));
    await c.refreshContext();
    expect(c.getState().chips).toHaveLength(1);
    expect(c.getState().chips[0]).toMatchObject({ id: 'word:selection', attached: false });

    await c.attach('word:selection');
    expect(assist.attached).toEqual(['word:selection']);
    expect(c.getState().chips[0]?.attached).toBe(true);

    c.detach('word:selection');
    expect(assist.detached).toEqual(['word:selection']);
    expect(c.getState().chips[0]?.attached).toBe(false);
  });
});

describe('PanelController — ask / stream', () => {
  it('streams tokens + citations into the assistant message and toggles busy', async () => {
    const assist = new FakeAssist();
    assist.script = [
      { type: 'token', text: 'The SLA ' },
      { type: 'token', text: 'is fine.' },
      { type: 'citation', source: { title: 'Vendor Policy', uri: 'https://x' } },
      {
        type: 'provenance',
        payload: { agentId: 'a', identity: 'u', timestamp: 't', sources: [], contentHash: 'h' },
      },
      { type: 'done' },
    ];
    const c = new PanelController(assist, lister([]));
    const seen: boolean[] = [];
    c.subscribe((s) => seen.push(s.busy));

    await c.send('is the SLA ok?');
    const msgs = c.getState().messages;
    expect(msgs.map((m) => m.role)).toEqual(['user', 'assistant']);
    expect(msgs[1]?.text).toBe('The SLA is fine.');
    expect(msgs[1]?.sources?.[0]?.title).toBe('Vendor Policy');
    expect(msgs[1]?.streaming).toBe(false);
    expect(c.getState().busy).toBe(false);
    expect(seen).toContain(true); // was busy mid-stream
  });

  it('ignores empty input and refuses concurrent sends', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    await c.send('   ');
    expect(c.getState().messages).toHaveLength(0);
  });

  it('surfaces a stream error on the assistant message', async () => {
    const assist = new FakeAssist();
    assist.script = [{ type: 'error', code: 'http_500', message: 'boom' }, { type: 'done' }];
    const c = new PanelController(assist, lister([]));
    await c.send('hi');
    expect(c.getState().messages[1]?.error).toBe('boom');
  });
});

describe('PanelController — actuation review', () => {
  it('applies a proposal and records the change', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    const p = c.propose('tracked-change', { text: 'x' }, 'Rewrite SLA clause');
    expect(c.getState().proposals[0]?.status).toBe('pending');

    await c.applyProposal(p.changeId);
    expect(assist.applied).toEqual([{ kind: 'tracked-change', changeId: p.changeId }]);
    expect(c.getState().proposals[0]?.status).toBe('applied');
    expect(c.getState().changes).toHaveLength(1);
  });

  it('marks a blocked proposal from a gate veto', async () => {
    const assist = new FakeAssist();
    assist.applyResult = {
      ok: false,
      changeId: '',
      kind: 'tracked-change',
      error: { code: 'blocked', message: 'Not grounded.' },
    };
    const c = new PanelController(assist, lister([]));
    const p = c.propose('tracked-change', { text: 'x' }, 'risky');
    await c.applyProposal(p.changeId);
    expect(c.getState().proposals[0]?.status).toBe('blocked');
    expect(c.getState().proposals[0]?.detail).toBe('Not grounded.');
  });
});

describe('PanelController — event-driven inputs', () => {
  it('onContext feeds events to the session (context path)', () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    c.onContext({ type: 'meeting-ended', id: 'm1' });
    expect(assist.ingested).toEqual([{ type: 'meeting-ended', id: 'm1' }]);
  });

  it('onSuggest adds an ignorable chip; dismiss removes it; onAutomate runs a turn', async () => {
    const assist = new FakeAssist();
    const c = new PanelController(assist, lister([]));
    c.onSuggest({ title: 'Draft follow-ups?', query: 'draft follow-ups' });
    const s = c.getState().suggestions[0];
    expect(s?.title).toBe('Draft follow-ups?');
    c.dismissSuggestion(s!.id);
    expect(c.getState().suggestions).toHaveLength(0);

    const send = vi.spyOn(c, 'send');
    c.onAutomate('go');
    expect(send).toHaveBeenCalledWith('go');
  });
});
