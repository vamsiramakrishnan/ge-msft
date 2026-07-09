// packages/runtime/src/docfs/skills-mount.test.ts
import { describe, expect, it } from 'vitest';
import { skillsMount } from './skills-mount.js';

function seeded() {
  return skillsMount({
    'm365-surface-commander/SKILL.md': '# Surface Commander\n\nTop-level playbook.\n',
    'm365-surface-commander/references/excel-semantics.md': '# Excel\n\nPivotTable notes.\n',
    'm365-surface-commander/references/word-semantics.md': '# Word\n\nTracked-change notes.\n',
    'm365-command-planner/SKILL.md': '# Planner\n',
  });
}

describe('skillsMount', () => {
  it('groups files under a common prefix into one dir entry at the root', async () => {
    const entries = await seeded().readdir('');
    expect(entries).toContainEqual({ name: 'm365-surface-commander', kind: 'dir' });
    expect(entries).toContainEqual({ name: 'm365-command-planner', kind: 'dir' });
    // No file is listed directly at the root — everything lives under a bundle directory.
    expect(entries.every((e) => e.kind === 'dir')).toBe(true);
  });

  it('lists files and nested dirs one level under a bundle', async () => {
    const entries = await seeded().readdir('m365-surface-commander');
    expect(entries).toContainEqual(expect.objectContaining({ name: 'SKILL.md', kind: 'file' }));
    expect(entries).toContainEqual({ name: 'references', kind: 'dir' });
  });

  it('lists files two levels deep', async () => {
    const entries = await seeded().readdir('m365-surface-commander/references');
    const names = entries.map((e) => e.name).sort();
    expect(names).toEqual(['excel-semantics.md', 'word-semantics.md']);
  });

  it('reads a file by its full relative path', async () => {
    const v = await seeded().readFile('m365-surface-commander/references/excel-semantics.md');
    expect(v?.text).toContain('PivotTable notes.');
  });

  it('returns null reading a path with no file (a directory, or nothing there)', async () => {
    expect(await seeded().readFile('m365-surface-commander')).toBeNull();
    expect(await seeded().readFile('nope.md')).toBeNull();
  });

  it('stat reports a file for a leaf path and a dir for a bundle prefix', async () => {
    const fs = seeded();
    expect(await fs.stat('m365-command-planner/SKILL.md')).toMatchObject({ kind: 'file' });
    expect(await fs.stat('m365-surface-commander')).toMatchObject({ kind: 'dir' });
    expect(await fs.stat('nope')).toBeNull();
  });

  it('searches file content and reports matching lines with their path', async () => {
    const hits = await seeded().search('', 'PivotTable');
    expect(hits).toEqual([
      {
        path: 'm365-surface-commander/references/excel-semantics.md',
        line: 3,
        text: 'PivotTable notes.',
      },
    ]);
  });

  it('search is scoped to files under the given path', async () => {
    const hits = await seeded().search('m365-command-planner', 'Planner');
    expect(hits).toHaveLength(1);
    expect(hits[0]?.path).toBe('m365-command-planner/SKILL.md');
  });

  it('search bounds the number of matches', async () => {
    const fs = skillsMount({
      'a.md': 'hit\nhit\nhit\n',
      'b.md': 'hit\nhit\nhit\n',
    });
    const hits = await fs.search('', 'hit', { max: 2 });
    expect(hits).toHaveLength(2);
  });

  it('returns [] searching for an empty pattern', async () => {
    expect(await seeded().search('', '')).toEqual([]);
  });

  it('degrades to an empty, harmless mount when given no files', async () => {
    const fs = skillsMount({});
    expect(await fs.readdir('')).toEqual([]);
    expect(await fs.readFile('anything')).toBeNull();
  });

  it('truncates readFile output to maxBytes', async () => {
    const fs = skillsMount({ 'big.md': 'x'.repeat(200) });
    const v = await fs.readFile('big.md', { maxBytes: 10 });
    expect(v?.truncated).toBe(true);
    expect(v?.bytes).toBeLessThanOrEqual(10);
  });
});
