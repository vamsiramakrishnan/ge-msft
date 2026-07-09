import { describe, expect, it } from 'vitest';
import {
  buildInsertArtifactProgram,
  tableToTsv,
  type InsertableArtifact,
} from './insert-artifact.js';

const TABLE: InsertableArtifact = {
  kind: 'markdown-table',
  title: 'Schedule',
  headers: ['Time', 'Monday', 'Tuesday'],
  rows: [
    ['08:00', 'India Sync', 'Music Lesson'],
    ['09:00', '**Deep Work**', '`Code Review`'],
  ],
};

describe('insert artifact command generation', () => {
  it('turns a markdown table into a rectangular TSV grid', () => {
    expect(tableToTsv(TABLE)).toBe(
      'Time\tMonday\tTuesday\n08:00\tIndia Sync\tMusic Lesson\n09:00\tDeep Work\tCode Review',
    );
  });

  it('builds one Excel grid command against the selected destination range', () => {
    const built = buildInsertArtifactProgram('excel', TABLE, { excelRange: "'Daily schedule'!C7" });
    expect(built).toEqual({
      ok: true,
      label: "Insert table into 'Daily schedule'!C7",
      program:
        'grid \'Daily schedule\'!C7 = "Time\\tMonday\\tTuesday\\n08:00\\tIndia Sync\\tMusic Lesson\\n09:00\\tDeep Work\\tCode Review"',
    });
  });

  it('fails closed for Excel when no destination range is known', () => {
    expect(buildInsertArtifactProgram('excel', TABLE)).toEqual({
      ok: false,
      reason: 'Select a destination range first.',
    });
  });

  it('maps table insertion to safe supported commands on other surfaces', () => {
    expect(buildInsertArtifactProgram('word', TABLE)).toMatchObject({
      ok: true,
      program:
        '/insert-text text="Time | Monday | Tuesday 08:00 | India Sync | Music Lesson 09:00 | Deep Work | Code Review"',
    });
    expect(buildInsertArtifactProgram('powerpoint', TABLE)).toMatchObject({
      ok: true,
      program:
        'slide "Schedule" "08:00 - India Sync - Music Lesson" "09:00 - Deep Work - Code Review"',
    });
    expect(buildInsertArtifactProgram('outlook', TABLE)).toMatchObject({
      ok: true,
      program:
        'compose "Schedule" "Time | Monday | Tuesday 08:00 | India Sync | Music Lesson 09:00 | Deep Work | Code Review"',
    });
  });

  it('lets code blocks be inserted through the same approval path', () => {
    const code: InsertableArtifact = { kind: 'code-block', code: 'total = revenue - cost' };
    expect(buildInsertArtifactProgram('teams', code)).toEqual({
      ok: true,
      label: 'Stage post',
      program: 'post "total = revenue - cost"',
    });
  });
});
