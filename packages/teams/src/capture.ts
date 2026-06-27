import type { Block } from '@ge/content';
import type { ResolvedContext } from '@ge/contracts';
import { native, toContext, type ToContextOptions } from '@ge/content';

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

/** Split a transcript window into non-empty, whitespace-normalized turn lines. Pure. */
export function transcriptToLines(transcript: string): string[] {
  return transcript
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter((line) => line.length > 0);
}

/** Cap on turn lines folded into the whole-transcript `read` snapshot. */
export const MAX_TRANSCRIPT_LINES = 60;

/** Cap on matching turn lines returned by a lazy `searchDocument` so a common term stays bounded. */
export const MAX_SEARCH_LINES = 8;

/**
 * The captured transcript window → the `Block[]` the surface-agnostic `buildDocStateSnapshot`
 * consumes for a whole-transcript `read` (ADR-0006). A transcript has no addressable sub-range, so
 * the "document" is the captured window: the meeting title becomes the outline heading and each turn
 * line a paragraph block, bounded by {@link MAX_TRANSCRIPT_LINES}. Anchored to a `transcript`
 * locator. Untrusted host content carried strictly as data.
 */
export function transcriptToDocStateBlocks(input: TranscriptInput): Block[] {
  const locator = 'transcript';
  const blocks: Block[] = [];
  if (input.meetingTitle?.trim()) blocks.push(native.heading(input.meetingTitle, 1, locator));
  for (const line of transcriptToLines(input.transcript).slice(0, MAX_TRANSCRIPT_LINES)) {
    blocks.push(native.paragraph(line, locator));
  }
  return blocks;
}

/**
 * Scan the transcript window for `query` (case-insensitive substring over turn lines) and return
 * the matching lines — labelled with the meeting metadata — as context via
 * {@link transcriptToContext}. Bounded to the first {@link MAX_SEARCH_LINES} matches. Pure. Empty
 * query / no match → `[]`.
 */
export function searchTranscript(input: TranscriptInput, query: string): ResolvedContext[] {
  const needle = query.trim().toLowerCase();
  if (!needle) return [];
  const matched: string[] = [];
  for (const line of transcriptToLines(input.transcript)) {
    if (line.toLowerCase().includes(needle)) {
      matched.push(line);
      if (matched.length >= MAX_SEARCH_LINES) break;
    }
  }
  if (matched.length === 0) return [];
  return transcriptToContext({
    ...(input.meetingTitle ? { meetingTitle: input.meetingTitle } : {}),
    ...(input.participants ? { participants: input.participants } : {}),
    transcript: matched.join('\n'),
  });
}
