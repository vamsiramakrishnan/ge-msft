import { describe, it, expect, vi } from 'vitest';
import { GraphSharedStore } from './shared-store.js';
import type { GraphClient } from './graph-client.js';

function fakeClient(): GraphClient & {
  listSharedFiles: ReturnType<typeof vi.fn>;
  getSharedFile: ReturnType<typeof vi.fn>;
  putSharedFile: ReturnType<typeof vi.fn>;
  deleteSharedFile: ReturnType<typeof vi.fn>;
} {
  return {
    listSharedFiles: vi.fn(async () => [{ name: 'notes.txt', size: 5 }]),
    getSharedFile: vi.fn(async () => 'hello'),
    putSharedFile: vi.fn(async () => undefined),
    deleteSharedFile: vi.fn(async () => undefined),
  } as unknown as GraphClient & {
    listSharedFiles: ReturnType<typeof vi.fn>;
    getSharedFile: ReturnType<typeof vi.fn>;
    putSharedFile: ReturnType<typeof vi.fn>;
    deleteSharedFile: ReturnType<typeof vi.fn>;
  };
}

describe('GraphSharedStore', () => {
  it('list() delegates to GraphClient.listSharedFiles', async () => {
    const client = fakeClient();
    const store = new GraphSharedStore(client);
    await expect(store.list()).resolves.toEqual([{ name: 'notes.txt', size: 5 }]);
    expect(client.listSharedFiles).toHaveBeenCalledOnce();
  });

  it('read() delegates to GraphClient.getSharedFile with the path', async () => {
    const client = fakeClient();
    const store = new GraphSharedStore(client);
    await expect(store.read('notes.txt')).resolves.toBe('hello');
    expect(client.getSharedFile).toHaveBeenCalledWith('notes.txt');
  });

  it('write() delegates to GraphClient.putSharedFile with the path and content', async () => {
    const client = fakeClient();
    const store = new GraphSharedStore(client);
    await store.write('notes.txt', 'updated');
    expect(client.putSharedFile).toHaveBeenCalledWith('notes.txt', 'updated');
  });

  it('remove() delegates to GraphClient.deleteSharedFile with the path', async () => {
    const client = fakeClient();
    const store = new GraphSharedStore(client);
    await store.remove('notes.txt');
    expect(client.deleteSharedFile).toHaveBeenCalledWith('notes.txt');
  });
});
