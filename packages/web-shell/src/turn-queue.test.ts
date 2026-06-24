import { describe, it, expect, vi } from 'vitest';
import type { ResolvedGrounding } from '@ge/gemini-client';
import { TurnQueue, type DrainHandlers } from './turn-queue.js';

const grounding = { sources: [] } as unknown as ResolvedGrounding;

const handlers = (): DrainHandlers & {
  ask: ReturnType<typeof vi.fn>;
  commands: ReturnType<typeof vi.fn>;
  skill: ReturnType<typeof vi.fn>;
} => ({
  ask: vi.fn(),
  commands: vi.fn(),
  skill: vi.fn(),
});

describe('TurnQueue', () => {
  it('starts empty and drains to nothing', () => {
    const q = new TurnQueue();
    const h = handlers();
    expect(q.queued).toBe(false);
    q.drain(h);
    expect(h.ask).not.toHaveBeenCalled();
    expect(h.commands).not.toHaveBeenCalled();
    expect(h.skill).not.toHaveBeenCalled();
  });

  it('preserves an ask turn through its own route (with grounding)', () => {
    const q = new TurnQueue();
    const h = handlers();
    q.enqueue({ mode: 'ask', query: 'why', grounding });
    expect(q.queued).toBe(true);
    q.drain(h);
    expect(h.ask).toHaveBeenCalledWith('why', grounding);
    expect(h.commands).not.toHaveBeenCalled();
  });

  it('preserves a commands turn — never downgraded to ask', () => {
    const q = new TurnQueue();
    const h = handlers();
    q.enqueue({ mode: 'commands', task: 'rewrite intro' });
    q.drain(h);
    expect(h.commands).toHaveBeenCalledWith('rewrite intro', undefined);
    expect(h.ask).not.toHaveBeenCalled();
    expect(h.skill).not.toHaveBeenCalled();
  });

  it('preserves a skill turn with its args', () => {
    const q = new TurnQueue();
    const h = handlers();
    q.enqueue({ mode: 'skill', name: 'summarize', args: { topic: 'q3' } });
    q.drain(h);
    expect(h.skill).toHaveBeenCalledWith('summarize', { topic: 'q3' });
  });

  it('is latest-wins: a second enqueue replaces the first', () => {
    const q = new TurnQueue();
    const h = handlers();
    q.enqueue({ mode: 'ask', query: 'first' });
    q.enqueue({ mode: 'commands', task: 'second' });
    q.drain(h);
    expect(h.ask).not.toHaveBeenCalled();
    expect(h.commands).toHaveBeenCalledWith('second', undefined);
  });

  it('clears the slot before dispatch so a re-enqueue during drain survives', () => {
    const q = new TurnQueue();
    const h = handlers();
    h.ask.mockImplementation(() => {
      // a dispatch that enqueues the next turn synchronously
      q.enqueue({ mode: 'commands', task: 'follow-up' });
    });
    q.enqueue({ mode: 'ask', query: 'go' });
    q.drain(h);
    expect(q.queued).toBe(true); // the re-enqueued turn was not clobbered
    q.drain(h);
    expect(h.commands).toHaveBeenCalledWith('follow-up', undefined);
  });

  it('drains only once per enqueue', () => {
    const q = new TurnQueue();
    const h = handlers();
    q.enqueue({ mode: 'ask', query: 'once' });
    q.drain(h);
    q.drain(h);
    expect(h.ask).toHaveBeenCalledTimes(1);
    expect(q.queued).toBe(false);
  });
});
