import { describe, it, expect } from 'vitest';
import { sessionStartEvent, sessionEndEvent, meetingEndedEvent } from './events.js';

describe('teams event builders (pure)', () => {
  it('builds a session-start tagged to the teams surface', () => {
    expect(sessionStartEvent()).toEqual({ type: 'session-start', surface: 'teams' });
  });

  it('builds a session-end tagged to the teams surface', () => {
    expect(sessionEndEvent()).toEqual({ type: 'session-end', surface: 'teams' });
  });

  it('builds a meeting-ended carrying its opaque id', () => {
    expect(meetingEndedEvent('19:meeting_abc')).toEqual({
      type: 'meeting-ended',
      id: '19:meeting_abc',
    });
  });
});
