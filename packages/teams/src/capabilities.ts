import type { CapabilityManifest } from '@ge/contracts';

/**
 * What the Teams surface can read and write. Read is the transcript window (meeting captions /
 * recent chat turns); write is a reviewable chat post / Adaptive Card. Static today;
 * `TeamsBridge.getCapabilities` could narrow it at runtime against TeamsJS `app.getContext()`
 * (e.g. drop `post-message` outside a chat/meeting frame, or drop `transcript` when captions are
 * unavailable).
 */
export const TEAMS_CAPABILITIES: CapabilityManifest = {
  surface: 'teams',
  contextKinds: ['transcript'],
  actuations: [
    {
      kind: 'post-message',
      surface: 'teams',
      title: 'Post to chat',
      description: 'Stage a reviewable, grounded chat post (or Adaptive Card) for the meeting.',
      reversible: true,
      appliesTo: ['transcript'],
    },
  ],
};
