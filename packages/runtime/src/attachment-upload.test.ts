import { describe, expect, it, vi } from 'vitest';
import type {
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  ResolvedContext,
} from '@ge/contracts';
import { StreamAssistClient } from '@ge/gemini-client';
import { AssistSession } from './assist-session.js';
import type { AttachmentBytes, DocBridge } from './bridge.js';

const cfg = { assistant: { project: 'p', location: 'eu', engine: 'e' }, identity: 'v.k@acme' };
const tokens = { getAccessToken: () => Promise.resolve('t') };
const unit = { connectors: [], surfaceContext: { kind: 'outlook' as const } };

const pdfRef: ContextRef = {
  id: 'outlook:attachment:a1',
  kind: 'attachment',
  surface: 'outlook',
  title: 'contract.pdf',
  mimeType: 'application/pdf',
  sizeBytes: 3,
};

function bridgeWith(opts: {
  readAttachment?: DocBridge['readAttachment'];
  parts: ResolvedContext[];
}) {
  const bridge: DocBridge = {
    surface: 'outlook',
    getCapabilities: (): CapabilityManifest => ({
      surface: 'outlook',
      contextKinds: ['attachment'],
      actuations: [],
    }),
    listContext: () => Promise.resolve([pdfRef]),
    resolveContext: () => Promise.resolve(opts.parts),
    actuate: (request: ActuationRequest): Promise<ActuationResult> =>
      Promise.resolve({ ok: true, changeId: request.changeId, kind: request.kind }),
    ...(opts.readAttachment ? { readAttachment: opts.readAttachment } : {}),
  };
  return bridge;
}

const fileValue = {
  as: 'file' as const,
  fileName: 'contract.pdf',
  mimeType: 'application/pdf',
  sizeBytes: 3,
  uploadable: true,
};
const filePart: ResolvedContext = { ref: pdfRef, value: fileValue };

describe('attachment upload on explicit attach', () => {
  it('reads the bytes through the bridge, uploads them, and records the fileId on the part', async () => {
    const readAttachment = vi.fn(
      (): Promise<AttachmentBytes> =>
        Promise.resolve({
          fileName: 'contract.pdf',
          mimeType: 'application/pdf',
          bytes: new Uint8Array([1, 2, 3]),
        }),
    );
    const client = new StreamAssistClient(tokens, cfg, vi.fn() as never);
    const upload = vi.spyOn(client, 'addContextFile').mockResolvedValue({
      fileId: 'file_1',
      fileName: 'contract.pdf',
      mimeType: 'application/pdf',
      byteSize: 3,
    });
    const session = new AssistSession(bridgeWith({ readAttachment, parts: [filePart] }), client, {
      unit,
    });

    await session.attachRef(pdfRef);

    expect(readAttachment).toHaveBeenCalledWith(pdfRef);
    expect(upload).toHaveBeenCalledTimes(1);
    expect(upload.mock.calls[0]![0]).toMatchObject({
      fileName: 'contract.pdf',
      mimeType: 'application/pdf',
    });
    const [attached] = session.context.list();
    expect(attached!.value).toMatchObject({ as: 'file', fileId: 'file_1', uploadable: true });
  });

  it('keeps the part by name with a note when the host cannot read attachments', async () => {
    const client = new StreamAssistClient(tokens, cfg, vi.fn() as never);
    const upload = vi.spyOn(client, 'addContextFile');
    const session = new AssistSession(bridgeWith({ parts: [filePart] }), client, { unit });
    await session.attachRef(pdfRef);
    expect(upload).not.toHaveBeenCalled();
    const [attached] = session.context.list();
    expect(attached!.value).toMatchObject({ as: 'file', uploadable: false });
    expect(attached!.value.as === 'file' && attached!.value.note).toContain('cannot read');
  });

  it('degrades to a note when the upload fails, and never fails the attach itself', async () => {
    const client = new StreamAssistClient(tokens, cfg, vi.fn() as never);
    vi.spyOn(client, 'addContextFile').mockRejectedValue(new Error('403 quota'));
    const bridge = bridgeWith({
      readAttachment: () =>
        Promise.resolve({
          fileName: 'contract.pdf',
          mimeType: 'application/pdf',
          bytes: new Uint8Array(1),
        }),
      parts: [filePart],
    });
    const session = new AssistSession(bridge, client, { unit });
    await expect(session.attachRef(pdfRef)).resolves.toBeUndefined();
    const [attached] = session.context.list();
    // The host/Graph error text stays out of the model-facing note.
    const note = attached!.value.as === 'file' ? attached!.value.note : undefined;
    expect(note).toContain('could not be attached');
    expect(note).not.toContain('403 quota');
    expect(attached!.value).not.toHaveProperty('fileId');
  });

  it('does not upload text parts or already-uploaded parts again', async () => {
    const client = new StreamAssistClient(tokens, cfg, vi.fn() as never);
    const upload = vi.spyOn(client, 'addContextFile');
    const done: ResolvedContext = { ...filePart, value: { ...fileValue, fileId: 'file_0' } };
    const text: ResolvedContext = {
      ref: { ...pdfRef, id: 'outlook:attachment:a1#0' },
      value: { as: 'text', text: 'File: notes.txt\n\nhello', mimeType: 'text/markdown' },
    };
    const session = new AssistSession(
      bridgeWith({
        readAttachment: () => Promise.reject(new Error('must not be called')),
        parts: [done, text],
      }),
      client,
      { unit },
    );
    await session.attachRef(pdfRef);
    expect(upload).not.toHaveBeenCalled();
    expect(session.context.size).toBe(2);
  });
});

describe('attachment upload reads through the attached ref', () => {
  it('passes the ref the user attached (with its host pointer), not the resolved part ref', async () => {
    const listed: ContextRef = {
      ...pdfRef,
      hostRef: { type: 'outlook.attachment', itemId: 'item-1', attachmentId: 'a1' },
    };
    const resolvedWithoutPointer: ResolvedContext = {
      ref: { ...pdfRef, id: `${pdfRef.id}#0` },
      value: fileValue,
    };
    const readAttachment = vi.fn(
      (): Promise<AttachmentBytes> =>
        Promise.resolve({
          fileName: 'contract.pdf',
          mimeType: 'application/pdf',
          bytes: new Uint8Array(1),
        }),
    );
    const client = new StreamAssistClient(tokens, cfg, vi.fn() as never);
    vi.spyOn(client, 'addContextFile').mockResolvedValue({
      fileId: 'file_9',
      fileName: 'contract.pdf',
      mimeType: 'application/pdf',
      byteSize: 1,
    });
    const session = new AssistSession(
      bridgeWith({ readAttachment, parts: [resolvedWithoutPointer] }),
      client,
      { unit },
    );
    await session.attachRef(listed);
    expect(readAttachment).toHaveBeenCalledWith(listed);
    expect(session.context.list()[0]!.value).toMatchObject({ as: 'file', fileId: 'file_9' });
  });
});
