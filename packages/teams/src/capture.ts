import type { ResolvedContext } from '@ge/contracts';
import { toContext, type ToContextOptions } from '@ge/content';

/**
 * Pure mapping from a Teams meeting/chat transcript window into grounding-ready context — no
 * TeamsJS here, so it's unit-testable. Like Outlook (and unlike Word/Excel), Teams has no native
 * block model the bridge can read: the transcript arrives as a **string** (a window of recent
 * turns), so we take @ge/content's string path. The `TeamsBridge` captures the transcript +
 * meeting metadata from TeamsJS and hands the plain values to this function, which labels them
 * (meeting title + participant roster) and normalizes the transcript body through `toContext`.
 *
 * The transcript is untrusted host content: it is carried as DATA (a labelled source string),
 * never interpreted as instructions.
 */
export interface TranscriptInput {
  meetingTitle?: string;
  transcript: string;
  participants?: string[];
}

export function transcriptToContext(
  input: TranscriptInput,
  opts: ToContextOptions = {},
): ResolvedContext[] {
  const sourceId = 'teams:transcript';
  const text = buildLabelledTranscript(input);
  return toContext(
    {
      sourceId,
      text,
      format: 'plain',
      surface: 'teams',
      ...(input.meetingTitle ? { title: input.meetingTitle } : {}),
    },
    opts,
  );
}

/**
 * Compose a single labelled source string from the meeting metadata and transcript window. The
 * `Meeting:` / `Participants:` prefix lines give the engine the same grounding cues a reader has,
 * and survive chunking as plain leading lines.
 */
function buildLabelledTranscript(input: TranscriptInput): string {
  const lines: string[] = [];
  if (input.meetingTitle) lines.push(`Meeting: ${input.meetingTitle}`);
  if (input.participants && input.participants.length > 0) {
    lines.push(`Participants: ${input.participants.join(', ')}`);
  }
  const header = lines.length > 0 ? `${lines.join('\n')}\n\n` : '';
  return `${header}${input.transcript}`;
}
