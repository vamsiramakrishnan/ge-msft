import { describe, expect, it } from 'vitest';

import { WorkspaceStore } from './workspace.js';

describe('WorkspaceStore', () => {
  it('stores bounded artifacts and clamps cat previews', () => {
    const store = new WorkspaceStore();
    store.save({
      name: 'long.txt',
      sourceLabel: 'literal',
      content: Array.from({ length: 300 }, (_, i) => `line ${i + 1}`).join('\n'),
    });

    const result = store.cat('long.txt', 500);

    expect(result.workspace).toBe('cat');
    if (result.workspace !== 'cat') throw new Error('expected cat result');
    expect(result.head).toBe(200);
    expect(result.preview.split('\n')).toHaveLength(200);
  });

  it('greps artifacts without re-reading the host', () => {
    const store = new WorkspaceStore();
    store.save({
      name: 'schedule.tsv',
      sourceLabel: "read 'Daily schedule'!B3:I53",
      content: 'Time\tMonday\n08:00\tDeep Work\n08:30\tManager Sync\n',
    });

    const result = store.grep('schedule.tsv', 'manager', 0);

    expect(result.workspace).toBe('grep');
    if (result.workspace !== 'grep') throw new Error('expected grep result');
    expect(result.matches).toEqual([{ line: 3, text: '08:30\tManager Sync' }]);
  });
});
