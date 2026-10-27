import { afterEach, describe, expect, it } from 'vitest';
import { OutlookBridge } from './outlook-bridge.js';
import { attachmentRefId } from './capture.js';

/**
 * Read-mode attachments through the real bridge against a minimal `Office` global: the item lists
 * `attachments[]` (Mailbox 1.1) and answers `getAttachmentContentAsync` (Mailbox 1.8) on the next
 * microtask, the way the host does.
 */
type Content = { format: string; content: string };
interface AttachmentSeed {
  id: string;
  name: string;
  attachmentType?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  content?: Content;
}

function install(attachments: AttachmentSeed[], opts: { mailboxMinor?: number } = {}): () => void {
  const g = globalThis as unknown as Record<string, unknown>;
  const prev = g.Office;
  const item = {
    itemId: 'item-1',
    subject: 'Q3 plan',
    from: { emailAddress: 'dana@contoso.com' },
    attachments: attachments.map(({ content: _c, ...details }) => details),
    body: {
      getAsync(_type: string, cb: (r: { status: string; value: string }) => void): void {
        queueMicrotask(() => cb({ status: 'succeeded', value: '<p>See attached.</p>' }));
      },
    },
    getAttachmentContentAsync(
      id: string,
      cb: (r: { status: string; value?: Content; error?: { message: string } }) => void,
    ): void {
      const found = attachments.find((a) => a.id === id);
      queueMicrotask(() =>
        found?.content
          ? cb({ status: 'succeeded', value: found.content })
          : cb({ status: 'failed', error: { message: 'no content' } }),
      );
    },
  };
  // Requirement sets compare by minor version (1.8 < 1.15), never as floats.
  const minor = opts.mailboxMinor ?? 15;
  g.Office = {
    CoercionType: { Html: 'html', Text: 'text' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    context: {
      mailbox: { item },
      requirements: {
        isSetSupported: (name: string, v: string) =>
          name === 'Mailbox' && Number(v.split('.')[1] ?? '0') <= minor,
      },
    },
  };
  return () => {
    g.Office = prev;
  };
}

const seeds: AttachmentSeed[] = [
  {
    id: 'a1',
    name: 'notes.txt',
    attachmentType: 'file',
    contentType: 'text/plain',
    size: 30,
    content: { format: 'base64', content: btoa('Decisions: ship in Q4.') },
  },
  { id: 'a2', name: 'logo.png', attachmentType: 'file', contentType: 'image/png', isInline: true },
  {
    id: 'a3',
    name: 'contract.pdf',
    attachmentType: 'file',
    contentType: 'application/pdf',
    size: 4,
    content: { format: 'base64', content: btoa('%PDF') },
  },
  {
    id: 'a4',
    name: 'FW: budget',
    attachmentType: 'item',
    content: { format: 'eml', content: 'Subject: budget\n\nApproved.' },
  },
];

describe('OutlookBridge — read-mode attachments', () => {
  let restore: (() => void) | undefined;
  afterEach(() => restore?.());

  it('lists attachments after the mail item, skipping inline images', async () => {
    restore = install(seeds);
    const refs = await new OutlookBridge().listContext();
    expect(refs.map((r) => r.id)).toEqual([
      'item-1',
      attachmentRefId('item-1', 'a1'),
      attachmentRefId('item-1', 'a3'),
      attachmentRefId('item-1', 'a4'),
    ]);
    expect(refs[1]).toMatchObject({ kind: 'attachment', title: 'notes.txt', sizeBytes: 30 });
  });

  it('resolves a text attachment to inline data and a PDF to an uploadable file part', async () => {
    restore = install(seeds);
    const bridge = new OutlookBridge();
    const refs = await bridge.listContext();
    const text = await bridge.resolveContext(refs[1]!);
    const joined = text.map((p) => (p.value.as === 'text' ? p.value.text : '')).join('\n');
    expect(joined).toContain('File: notes.txt');
    expect(joined).toContain('Decisions: ship in Q4.');
    const [pdf] = await bridge.resolveContext(refs[2]!);
    expect(pdf!.value).toMatchObject({ as: 'file', uploadable: true, mimeType: 'application/pdf' });
    const [eml] = await bridge.resolveContext(refs[3]!);
    expect(eml!.value.as === 'text' && eml!.value.text).toContain('Approved.');
  });

  it("reads a file attachment's bytes for the session upload", async () => {
    restore = install(seeds);
    const bridge = new OutlookBridge();
    const refs = await bridge.listContext();
    const raw = await bridge.readAttachment(refs[2]!);
    expect(raw).toMatchObject({ fileName: 'contract.pdf', mimeType: 'application/pdf' });
    expect(new TextDecoder().decode(raw.bytes)).toBe('%PDF');
    await expect(bridge.readAttachment(refs[3]!)).rejects.toThrow('cannot read');
  });

  it('grounds attachments by name only on a host without Mailbox 1.8', async () => {
    restore = install(seeds, { mailboxMinor: 7 });
    const bridge = new OutlookBridge();
    const refs = await bridge.listContext();
    const [part] = await bridge.resolveContext(refs[1]!);
    expect(part!.value).toMatchObject({ as: 'file', uploadable: false });
    await expect(bridge.readAttachment(refs[2]!)).rejects.toThrow('cannot read');
  });

  it('refuses an attachment ref from another mail item', async () => {
    restore = install(seeds);
    const bridge = new OutlookBridge();
    const refs = await bridge.listContext();
    const stale = {
      ...refs[1]!,
      hostRef: { type: 'outlook.attachment' as const, itemId: 'item-0', attachmentId: 'a1' },
    };
    await expect(bridge.resolveContext(stale)).rejects.toThrow('no longer active');
  });
});

describe('OutlookBridge — resolved parts keep the host pointer', () => {
  let restore: (() => void) | undefined;
  afterEach(() => restore?.());

  it('lets the runtime read bytes through a resolved file part as well as the listed ref', async () => {
    restore = install(seeds);
    const bridge = new OutlookBridge();
    const refs = await bridge.listContext();
    const [pdf] = await bridge.resolveContext(refs[2]!);
    expect(pdf!.ref.hostRef).toEqual({
      type: 'outlook.attachment',
      itemId: 'item-1',
      attachmentId: 'a3',
    });
    const raw = await bridge.readAttachment(pdf!.ref);
    expect(raw.fileName).toBe('contract.pdf');
  });
});
