import { describe, expect, it, vi } from 'vitest';
import { currentTeamsChatId, graphSharedFiles } from './teams-shared-files.js';

describe('graphSharedFiles', () => {
  it('lists chat files through Graph and reads only files it listed', async () => {
    const graph = {
      listChatFiles: vi.fn(() =>
        Promise.resolve([
          {
            id: 'i1',
            driveId: 'd1',
            name: 'a.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 5,
            sharedBy: 'Dana',
          },
        ]),
      ),
      getDriveItemContent: vi.fn(() =>
        Promise.resolve({ bytes: new Uint8Array([1]), mimeType: 'application/pdf' }),
      ),
    };
    const source = graphSharedFiles(graph, '19:chat', { maxBytes: 100 });
    await expect(source.read('i1')).rejects.toThrow('not listed');
    expect(await source.list()).toEqual([
      { id: 'i1', name: 'a.pdf', mimeType: 'application/pdf', sizeBytes: 5, sharedBy: 'Dana' },
    ]);
    expect(graph.listChatFiles).toHaveBeenCalledWith('19:chat');
    const read = await source.read('i1');
    expect([...read.bytes]).toEqual([1]);
    expect(graph.getDriveItemContent).toHaveBeenCalledWith('i1', 'd1', { maxBytes: 100 });
  });
});

describe('currentTeamsChatId', () => {
  it('reads the chat id from TeamsJS app.getContext, tolerating absence and failure', async () => {
    expect(await currentTeamsChatId(undefined)).toBeUndefined();
    expect(await currentTeamsChatId({})).toBeUndefined();
    expect(
      await currentTeamsChatId({
        app: { getContext: () => Promise.resolve({ chat: { id: '19:c' } }) },
      }),
    ).toBe('19:c');
    expect(
      await currentTeamsChatId({ app: { getContext: () => Promise.reject(new Error('no')) } }),
    ).toBeUndefined();
  });
});
