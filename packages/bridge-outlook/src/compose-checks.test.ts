import { afterEach, describe, expect, it, vi } from 'vitest';
import { TriggerRegistry } from '@ge/triggers';
import { captureComposeFacts, checkComposeFacts, type ComposeFacts } from './compose-checks.js';
import { createMessageSendHandler } from './on-send.js';

const facts = (): ComposeFacts => ({
  instanceId: 'compose-fixture',
  itemId: 'draft-1',
  subject: 'Board update',
  body: 'Revenue: 120.00',
  to: ['pat@example.com'],
  cc: [],
  bcc: [],
  attachments: [{ id: 'a1', name: 'report.pdf', isInline: false }],
});
afterEach(() => {
  vi.unstubAllGlobals();
});

function install(source = facts(), configure?: (item: Record<string, unknown>) => void) {
  const ok = (value: unknown) => ({ status: 'succeeded', value });
  const field = (value: unknown) => ({
    getAsync: (cb: (result: unknown) => void) => cb(ok(value)),
  });
  const item: Record<string, unknown> = {
    itemId: source.itemId,
    subject: field(source.subject),
    body: { getAsync: (_type: unknown, cb: (result: unknown) => void) => cb(ok(source.body)) },
    to: field(source.to.map((emailAddress) => ({ emailAddress }))),
    cc: field(source.cc.map((emailAddress) => ({ emailAddress }))),
    bcc: field(source.bcc.map((emailAddress) => ({ emailAddress }))),
    getAttachmentsAsync: (cb: (result: unknown) => void) => cb(ok(source.attachments)),
  };
  configure?.(item);
  const mailbox = { item };
  vi.stubGlobal('Office', {
    context: { mailbox, requirements: { isSetSupported: () => true } },
    AsyncResultStatus: { Succeeded: 'succeeded' },
    CoercionType: { Text: 'text' },
  });
  return mailbox;
}

describe('complete compose facts and deterministic policy', () => {
  it('reads subject/body, all recipient fields and only attachment metadata from native callbacks', async () => {
    install();
    expect(await captureComposeFacts()).toEqual({
      ...facts(),
      instanceId: expect.stringMatching(/^compose_\d+$/),
    });
  });
  it('requires the complete source instead of treating absent recipient APIs as no recipients', async () => {
    install(facts(), (item) => {
      delete item.bcc;
    });
    await expect(captureComposeFacts()).rejects.toThrow('compose APIs');
  });
  it('rejects an older requirement set even if proxy members exist', async () => {
    install();
    vi.stubGlobal('Office', {
      context: { mailbox: Office.context.mailbox, requirements: { isSetSupported: () => false } },
    });
    await expect(captureComposeFacts()).rejects.toThrow('Mailbox 1.8');
  });
  it('rejects a body above the UTF-8 byte budget without certifying a clipped prefix', async () => {
    install({ ...facts(), body: '😀'.repeat(17000) });
    await expect(captureComposeFacts()).rejects.toThrow('budget');
  });
  it('rejects a switch to another item during asynchronous callbacks', async () => {
    let complete: ((result: unknown) => void) | undefined;
    const mailbox = install(facts(), (item) => {
      item.subject = {
        getAsync: (cb: (result: unknown) => void) => {
          complete = cb;
        },
      };
    });
    const capture = captureComposeFacts();
    mailbox.item = {};
    complete!({ status: 'succeeded', value: 'Board update' });
    await expect(capture).rejects.toThrow('active draft changed');
  });
  it('checks Bcc and exact recipient domains; a permitted suffix does not permit subdomains', () => {
    expect(
      checkComposeFacts(
        { ...facts(), bcc: ['x@external.com'] },
        { allowedRecipientDomains: ['example.com'] },
      ).allowEvent,
    ).toBe(false);
    expect(
      checkComposeFacts(
        { ...facts(), to: ['x@child.example.com'] },
        { allowedRecipientDomains: ['example.com'] },
      ).allowEvent,
    ).toBe(false);
    expect(
      checkComposeFacts(facts(), { allowedRecipientDomains: ['EXAMPLE.COM'] }).allowEvent,
    ).toBe(true);
  });
  it('checks required subject, recipients and actual non-inline attachment filenames', () => {
    expect(
      checkComposeFacts({ ...facts(), subject: ' ' }, { requireSubject: true }).allowEvent,
    ).toBe(false);
    expect(checkComposeFacts({ ...facts(), to: [] }, { requireRecipient: true }).allowEvent).toBe(
      false,
    );
    expect(checkComposeFacts(facts(), { requiredAttachments: ['report.pdf'] }).allowEvent).toBe(
      true,
    );
    expect(
      checkComposeFacts(
        { ...facts(), attachments: [{ id: 'a', name: 'report.pdf', isInline: true }] },
        { requiredAttachments: ['report.pdf'] },
      ).allowEvent,
    ).toBe(false);
  });
});

describe('compose checks in the send-event boundary', () => {
  it('allows only unchanged facts after the registered checks finish', async () => {
    const capture = vi.fn(async () => facts());
    const completed = vi.fn();
    await createMessageSendHandler(new TriggerRegistry(), {
      composePolicy: { requireRecipient: true },
      captureCompose: capture,
    })({ completed });
    expect(capture).toHaveBeenCalledTimes(2);
    expect(completed).toHaveBeenCalledTimes(1);
    expect(completed).toHaveBeenCalledWith({ allowEvent: true });
  });
  it('blocks changed recipient/body/attachment facts without claiming the old check still applies', async () => {
    const capture = vi
      .fn()
      .mockResolvedValueOnce(facts())
      .mockResolvedValueOnce({ ...facts(), bcc: ['x@external.com'] });
    const completed = vi.fn();
    await createMessageSendHandler(new TriggerRegistry(), {
      composePolicy: { requireRecipient: true },
      captureCompose: capture,
    })({ completed });
    expect(completed).toHaveBeenCalledTimes(1);
    expect(completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage: 'The draft changed during send checks. Review it and send again.',
    });
  });
  it('fails closed on missing/failed host capture and completes once', async () => {
    const completed = vi.fn();
    await createMessageSendHandler(new TriggerRegistry(), {
      composePolicy: { requireSubject: true },
      captureCompose: async () => {
        throw new Error('secret body');
      },
    })({ completed });
    expect(completed).toHaveBeenCalledTimes(1);
    expect(completed).toHaveBeenCalledWith({
      allowEvent: false,
      errorMessage: 'The send checks could not complete. Try again before sending this message.',
    });
  });
});
