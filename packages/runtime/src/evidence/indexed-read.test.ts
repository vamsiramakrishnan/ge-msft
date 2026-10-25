import { describe, expect, it, vi } from 'vitest';
import { EvidenceWorkspace } from './workspace.js';
import { MemoryEvidenceStore } from './store.js';
import { evidenceMount } from './mount.js';

describe('indexed evidence reads', () => {
  it('reads one hashed page without materializing the complete source snapshot', async () => {
    const store = new MemoryEvidenceStore();
    const workspace = new EvidenceWorkspace(store);
    const receipt = await workspace.capture({
      id: 'large',
      readPage: async ({ cursor }) => {
        const index = Number(cursor ?? 0);
        return {
          scope: 'body',
          units: [{ id: 'unit:' + index, text: 'page ' + index }],
          limitations: [],
          ...(index < 99 ? { nextCursor: String(index + 1) } : {}),
        };
      },
    });
    const full = vi.spyOn(store, 'get');
    const page = vi.spyOn(store, 'getPage');
    expect((await workspace.readText(receipt.id, 'unit:75')).text).toBe('page 75');
    expect(full).not.toHaveBeenCalled();
    expect(page).toHaveBeenCalledTimes(1);
    expect(page).toHaveBeenCalledWith(receipt.id, 75);
    const file = await evidenceMount(workspace).readFile(`${receipt.id}/units/0/75.json`);
    expect(file?.text).toContain('page 75');
    expect(full).not.toHaveBeenCalled();
    page.mockClear();
    const query = await workspace.query(receipt.id, { kind: 'grep', literal: 'page 99' });
    expect(query.complete).toBe(false);
    expect(query.nextCursor).toBeDefined();
    expect(page.mock.calls.length).toBeLessThanOrEqual(8);
    await workspace.release(receipt.id);
    await expect(workspace.read(receipt.id, 'unit:75')).rejects.toThrow('unavailable');
  });
  it('rejects tampered pages and wrong index membership from persistent adapters', async () => {
    const store = new MemoryEvidenceStore();
    const workspace = new EvidenceWorkspace(store);
    const { id } = await workspace.capture({
      id: 'source',
      readPage: async () => ({
        scope: 'body',
        units: [{ id: 'a', text: 'original' }],
        limitations: [],
      }),
    });
    vi.spyOn(store, 'getPage').mockResolvedValue({
      scope: 'body',
      units: [{ id: 'a', text: 'tampered' }],
      limitations: [],
    });
    await expect(workspace.read(id, 'a')).rejects.toThrow('integrity');
  });
});
