import { describe, it, expect } from 'vitest';
import type { HostEvent } from '@ge/triggers';
import { TeamsBridge, type TeamsJsLike } from './teams-bridge.js';

describe('TeamsBridge context capture', () => {
  it('lists no context before a transcript is captured', async () => {
    const bridge = new TeamsBridge();
    expect(await bridge.listContext()).toEqual([]);
  });

  it('lists the transcript window once captured', async () => {
    const bridge = new TeamsBridge({
      transcript: { meetingTitle: 'Sync', transcript: 'Pat: hi.' },
    });
    const refs = await bridge.listContext();
    expect(refs).toHaveLength(1);
    expect(refs[0]?.kind).toBe('transcript');
    expect(refs[0]?.surface).toBe('teams');
  });

  it('reflects a transcript updated via setTranscript', async () => {
    const bridge = new TeamsBridge();
    bridge.setTranscript({ transcript: 'Now there are turns.' });
    const refs = await bridge.listContext();
    expect(refs).toHaveLength(1);
    const ctx = await bridge.resolveContext(refs[0]!);
    expect(ctx.length).toBeGreaterThan(0);
  });
});

describe('TeamsBridge actuate (post-message)', () => {
  it('rejects an unsupported kind', async () => {
    const bridge = new TeamsBridge();
    const res = await bridge.actuate({
      changeId: 'c1',
      kind: 'insert-text',
      surface: 'teams',
      params: { text: 'x' },
    });
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('unsupported');
  });

  it('errors when post-message has no content', async () => {
    const bridge = new TeamsBridge();
    const res = await bridge.actuate({
      changeId: 'c2',
      kind: 'post-message',
      surface: 'teams',
      params: {},
    });
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('no_content');
  });

  it('degrades to a panel item when no host compose path exists', async () => {
    const bridge = new TeamsBridge();
    const res = await bridge.actuate({
      changeId: 'c3',
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Action items.' },
    });
    expect(res.ok).toBe(true);
    expect(res.degraded).toBe(true);
    expect(res.location).toBe('panel');
  });

  it('stages a reviewable post through the host compose path', async () => {
    const sent: unknown[] = [];
    const teams: TeamsJsLike = {
      chat: { openConversation: (req) => sent.push(req) },
    };
    const bridge = new TeamsBridge({ teams });
    const res = await bridge.actuate({
      changeId: 'c4',
      kind: 'post-message',
      surface: 'teams',
      params: { text: 'Summary', html: '<card/>' },
    });
    expect(res.ok).toBe(true);
    expect(res.location).toBe('chat-compose');
    expect(sent).toEqual([{ message: 'Summary', card: '<card/>' }]);
  });
});

describe('TeamsBridge watch', () => {
  it('emits session-start on registration and session-end on unsubscribe', () => {
    const events: HostEvent[] = [];
    const bridge = new TeamsBridge();
    const unsub = bridge.watch((e) => events.push(e));
    expect(events).toEqual([{ type: 'session-start', surface: 'teams' }]);
    unsub();
    expect(events).toEqual([
      { type: 'session-start', surface: 'teams' },
      { type: 'session-end', surface: 'teams' },
    ]);
  });

  it('emits meeting-ended when TeamsJS raises its end handler', () => {
    const events: HostEvent[] = [];
    let endHandler: (() => void) | undefined;
    const teams: TeamsJsLike = {
      meeting: { registerMeetingEndHandler: (h) => (endHandler = h) },
    };
    const bridge = new TeamsBridge({ teams, meetingId: '19:abc' });
    bridge.watch((e) => events.push(e));
    endHandler?.();
    expect(events).toContainEqual({ type: 'meeting-ended', id: '19:abc' });
  });

  it('never throws when TeamsJS registration explodes; returns a clean unsubscribe', () => {
    const teams: TeamsJsLike = {
      meeting: {
        registerMeetingEndHandler: () => {
          throw new Error('wrong frame');
        },
      },
    };
    const bridge = new TeamsBridge({ teams });
    expect(() => {
      const unsub = bridge.watch(() => {});
      unsub();
    }).not.toThrow();
  });
});
