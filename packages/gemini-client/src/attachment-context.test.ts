import { describe, expect, it } from 'vitest';
import type { AssistRequest, ResolvedContext } from '@ge/contracts';
import { buildStreamAssistRequest } from './stream-assist.js';
import { contextValueToQueryPart, fileManifestLine } from './session-context.js';
import type { GeminiClientConfig } from './config.js';

const cfg: GeminiClientConfig = {
  assistant: { project: 'p', location: 'eu', engine: 'e' },
  identity: 'v.k@acme',
};
const req: AssistRequest = {
  intent: 'ask',
  query: 'Summarize the attachment.',
  unit: { connectors: [], surfaceContext: { kind: 'outlook' } },
};

const uploaded: ResolvedContext = {
  ref: { id: 'outlook:attachment:a1', kind: 'attachment', surface: 'outlook', title: 'deck.pptx' },
  value: {
    as: 'file',
    fileName: 'deck.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    sizeBytes: 2_500_000,
    uploadable: true,
    fileId: 'file_123',
  },
};
const byNameOnly: ResolvedContext = {
  ref: { id: 'outlook:attachment:a2', kind: 'attachment', surface: 'outlook', title: 'photo.png' },
  value: {
    as: 'file',
    fileName: 'photo.png',
    mimeType: 'image/png',
    uploadable: false,
    note: 'image/png is not readable here; content omitted',
  },
};

describe('file context values', () => {
  it('names the file in a data part and says whether its content is available', () => {
    expect(fileManifestLine(uploaded.value as never)).toBe(
      '[Attached file: deck.pptx (application/vnd.openxmlformats-officedocument.presentationml.presentation, 2.4 MB) — attached as a session context file]',
    );
    expect(contextValueToQueryPart(byNameOnly.value)).toEqual({
      text: '[Attached file: photo.png (image/png) — image/png is not readable here; content omitted]',
    });
  });

  it('carries uploaded attachment fileIds on a session request, deduplicated with grounding', () => {
    const body = buildStreamAssistRequest(req, cfg, 'sess_1', [uploaded, byNameOnly], 'default', {
      fileIds: ['file_123', 'file_999'],
    });
    expect(body.fileIds).toEqual(['file_123', 'file_999']);
    const parts = (body.query as { parts: Array<{ text?: string }> }).parts;
    expect(parts.some((p) => p.text?.includes('deck.pptx'))).toBe(true);
    expect(parts.some((p) => p.text?.includes('photo.png'))).toBe(true);
  });

  it('drops attached fileIds on a sessionless request instead of failing it', () => {
    const body = buildStreamAssistRequest(req, cfg, undefined, [uploaded], 'default', undefined, {
      isSessionLess: true,
    });
    expect(body.fileIds).toBeUndefined();
    expect(body.isSessionLess).toBe(true);
    const parts = (body.query as { parts: Array<{ text?: string }> }).parts;
    expect(parts.some((p) => p.text?.includes('deck.pptx'))).toBe(true);
  });
});
