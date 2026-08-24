import { afterEach, describe, expect, it } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationKind,
  type ActuationParams,
  type ActuationRequest,
  type ContextRef,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { MAX_ATTACHMENT_BASE64_CHARS } from './actuate-plan.js';
import { HANDLED_ACTUATIONS, OutlookBridge } from './outlook-bridge.js';

/**
 * Self-contained, in-memory **Outlook mailbox simulator** for the bridge-outlook package. It models
 * ONLY the slice of the Office.js mailbox object model {@link OutlookBridge} drives: the single
 * active `Office.context.mailbox.item` (read-mode statics AND the compose-mode write callbacks —
 * `Recipients.setAsync/addAsync`, `Subject.setAsync`, `Body.setAsync`,
 * `addFileAttachmentFromBase64Async`/`addFileAttachmentAsync`/`addItemAttachmentAsync`), the
 * compose-form launchers (`displayReplyForm` / `displayNewMessageForm`), and the `ItemChanged`
 * event registration. Reads are callback-based exactly like the host; writes record into the seed
 * so a test can assert the reviewable form actually opened with the right payload and that draft
 * edits never send.
 *
 * Kept local to this package (mirrors the fake-excel harness shape) so it adds no cross-package
 * dependency. The mail body/subject are UNTRUSTED — these tests assert the bridge never auto-sends
 * and that the reviewable-form/draft boundary is the only write path.
 */

/* ───────────────────────────── seed + recorders ────────────────────────── */

interface ReplyFormCall {
  htmlBody?: string;
}
interface NewMessageFormCall {
  subject?: string;
  htmlBody?: string;
  toRecipients?: string[];
}

interface RecipientWrite {
  mode: 'set' | 'add';
  field: string;
  addresses: string[];
}
interface SubjectSet {
  subject: string;
}
interface BodySet {
  data: string;
  coercionType: unknown;
}
interface AttachmentAdded {
  transport: 'uri' | 'base64' | 'item';
  value: string;
  name: string;
  isInline: boolean;
}

/** Compose-mode draft state behind the active item (drives the write-capable host objects). */
interface DraftSeed {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  /** Field whose setAsync/addAsync invokes its callback with a failed status. */
  recipientsFailsOn?: 'to' | 'cc' | 'bcc';
  subjectSetFails?: boolean;
  bodySetFails?: boolean;
  attachmentAddFails?: boolean;
  /** Strip every compose write API → the item looks like a READ-mode message. */
  readOnly?: boolean;
}

interface MailItemSeed {
  itemId?: string;
  subject?: string;
  from?: { emailAddress?: string };
  /** Raw body string the fake `body.getAsync` returns. */
  body: string;
  /** When true, `body.getAsync` invokes its callback with a failed status (host read error). */
  bodyFails?: boolean;
  /** When present (and not readOnly), exposes the compose-mode write APIs on the item. */
  draft?: DraftSeed;
}

interface OutlookHandle {
  type: unknown;
  handler: () => void;
}

interface Installed {
  replyForms: ReplyFormCall[];
  newMessageForms: NewMessageFormCall[];
  openedMessages: string[];
  handlers: OutlookHandle[];
  recipientWrites: RecipientWrite[];
  subjectSets: SubjectSet[];
  bodySets: BodySet[];
  attachmentsAdded: AttachmentAdded[];
  /** Swap the active item (or clear it) to simulate ItemChanged / a switch to a draft. */
  setItem(seed: MailItemSeed | undefined): void;
  /** Fire the registered ItemChanged handler(s). */
  fireItemChanged(): void;
  restore(): void;
}

/**
 * Install a fake `globalThis.Office`. `opts.noMailbox` drops the mailbox entirely (wrong host /
 * pane not in a mailbox); `opts.noNewMessageForm` keeps the mailbox but removes
 * `displayNewMessageForm` (compose unavailable on this host); `opts.throwOnAdd` makes
 * `addHandlerAsync` throw; `opts.throwOnRemove` makes teardown throw.
 */
function installOutlook(
  initial: MailItemSeed | undefined,
  opts: {
    noMailbox?: boolean;
    noNewMessageForm?: boolean;
    throwOnAdd?: boolean;
    throwOnRemove?: boolean;
  } = {},
): Installed {
  const replyForms: ReplyFormCall[] = [];
  const newMessageForms: NewMessageFormCall[] = [];
  const openedMessages: string[] = [];
  const handlers: OutlookHandle[] = [];
  const recipientWrites: RecipientWrite[] = [];
  const subjectSets: SubjectSet[] = [];
  const bodySets: BodySet[] = [];
  const attachmentsAdded: AttachmentAdded[] = [];
  let current: MailItemSeed | undefined = initial;

  type ResultCb<T> = (r: { status: string; value?: T; error?: { message?: string } }) => void;
  const ok = <T>(value: T) => ({ status: 'succeeded', value });
  const fail = (message: string) => ({ status: 'failed', error: { message } });

  const buildItem = (seed: MailItemSeed | undefined): unknown => {
    if (!seed) return undefined;
    const d = seed.draft;
    const compose = d !== undefined && !d.readOnly;

    const recipientField = (field: 'to' | 'cc' | 'bcc') => ({
      setAsync(list: string[], cb: ResultCb<void>): void {
        recipientWrites.push({ mode: 'set', field, addresses: list });
        if (d?.recipientsFailsOn === field) {
          cb(fail('recipient write failed'));
        } else {
          if (d) d[field] = [...list];
          cb(ok(undefined));
        }
      },
      addAsync(list: string[], cb: ResultCb<void>): void {
        recipientWrites.push({ mode: 'add', field, addresses: list });
        if (d?.recipientsFailsOn === field) {
          cb(fail('recipient write failed'));
        } else {
          if (d) d[field] = [...(d[field] ?? []), ...list];
          cb(ok(undefined));
        }
      },
    });

    const attachmentApi =
      (transport: AttachmentAdded['transport']) =>
      (
        value: string,
        name: string,
        options: { isInline?: boolean },
        cb: ResultCb<string>,
      ): void => {
        attachmentsAdded.push({
          transport,
          value,
          name,
          isInline: options?.isInline === true,
        });
        if (d?.attachmentAddFails) {
          cb(fail('attachment upload failed'));
        } else {
          cb(ok(`att-${transport}-1`));
        }
      };

    return {
      ...(seed.itemId !== undefined ? { itemId: seed.itemId } : {}),
      // Compose mode exposes Subject as a write-capable object; read mode as a static string.
      ...(compose
        ? {
            subject: {
              getAsync(cb: ResultCb<string>): void {
                cb(ok(seed.subject ?? ''));
              },
              setAsync(value: string, cb: ResultCb<void>): void {
                subjectSets.push({ subject: value });
                if (d.subjectSetFails) cb(fail('subject write failed'));
                else cb(ok(undefined));
              },
            },
          }
        : seed.subject !== undefined
          ? { subject: seed.subject }
          : {}),
      ...(seed.from !== undefined ? { from: seed.from } : {}),
      ...(compose && d.to !== undefined ? { to: recipientField('to') } : {}),
      ...(compose && d.cc !== undefined ? { cc: recipientField('cc') } : {}),
      ...(compose && d.bcc !== undefined ? { bcc: recipientField('bcc') } : {}),
      body: {
        getAsync(
          _coercion: unknown,
          cb: (r: { status: string; value?: string; error?: { message?: string } }) => void,
        ): void {
          if (seed.bodyFails) {
            cb({ status: 'failed', error: { message: 'body read failed' } });
          } else {
            cb({ status: 'succeeded', value: seed.body });
          }
        },
        ...(compose
          ? {
              setAsync(
                data: string,
                options: { coercionType?: unknown },
                cb: ResultCb<void>,
              ): void {
                bodySets.push({ data, coercionType: options?.coercionType });
                if (d.bodySetFails) cb(fail('body write failed'));
                else cb(ok(undefined));
              },
            }
          : {}),
      },
      ...(compose
        ? {
            addFileAttachmentAsync: attachmentApi('uri'),
            addFileAttachmentFromBase64Async: attachmentApi('base64'),
            addItemAttachmentAsync: attachmentApi('item'),
          }
        : {}),
      displayReplyForm(arg: ReplyFormCall): void {
        replyForms.push(arg);
      },
    };
  };

  const mailbox = opts.noMailbox
    ? undefined
    : {
        get item(): unknown {
          return buildItem(current);
        },
        ...(opts.noNewMessageForm
          ? {}
          : {
              displayNewMessageForm(arg: NewMessageFormCall): void {
                newMessageForms.push(arg);
              },
            }),
        displayMessageForm(itemId: string): void {
          openedMessages.push(itemId);
        },
        addHandlerAsync(type: unknown, handler: () => void): void {
          if (opts.throwOnAdd) throw new Error('addHandlerAsync failed');
          handlers.push({ type, handler });
        },
        removeHandlerAsync(type: unknown, handler: () => void): void {
          if (opts.throwOnRemove) throw new Error('removeHandlerAsync failed');
          const i = handlers.findIndex((h) => h.type === type && h.handler === handler);
          if (i >= 0) handlers.splice(i, 1);
        },
      };

  const office = {
    context: { mailbox },
    CoercionType: { Html: 'html', Text: 'text' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
  };

  const g = globalThis as unknown as Record<string, unknown>;
  const prevOffice = g.Office;
  g.Office = office;

  return {
    replyForms,
    newMessageForms,
    openedMessages,
    handlers,
    recipientWrites,
    subjectSets,
    bodySets,
    attachmentsAdded,
    setItem(seed) {
      current = seed;
    },
    fireItemChanged() {
      for (const h of [...handlers]) h.handler();
    },
    restore() {
      g.Office = prevOffice;
    },
  };
}

/* ───────────────────────────── fixtures ────────────────────────────────── */

function readMail(): MailItemSeed {
  return {
    itemId: 'AAMk-9',
    subject: 'SLA concerns',
    from: { emailAddress: 'pat@acme.com' },
    body: '<p>We are at <strong>99.5%</strong> availability.</p><p>Can we raise the SLA to 99.9%?</p>',
  };
}

function reply(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'reply-mail', surface: 'outlook', params };
}
function compose(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'create-mail', surface: 'outlook', params };
}
function act(kind: ActuationKind, params: ActuationParams, id = 'chg-x'): ActuationRequest {
  return { changeId: asChangeId(id), kind, surface: 'outlook', params };
}

/** A compose-mode draft item: functional To/Cc/Bcc (empty), editable subject/body, attachments. */
function draftMail(overrides: DraftSeed = {}): MailItemSeed {
  return {
    subject: 'Q3 follow-up',
    body: '<p>Current draft body.</p>',
    draft: { to: [], cc: [], bcc: [], ...overrides },
  };
}

let active: Installed | undefined;
afterEach(() => {
  active?.restore();
  active = undefined;
});

/* ───────────────────────────── tests ───────────────────────────────────── */

describe('OutlookBridge basics', () => {
  it('exposes the outlook surface and the manifest capabilities', () => {
    const bridge = new OutlookBridge();
    expect(bridge.surface).toBe('outlook');
    const caps = bridge.getCapabilities();
    expect(caps.surface).toBe('outlook');
    expect(new Set(caps.actuations.map((a) => a.kind))).toEqual(new Set(HANDLED_ACTUATIONS));
  });
});

describe('OutlookBridge.listContext (host wiring)', () => {
  it('returns a single live mail-item ref labelled with the subject', async () => {
    active = installOutlook(readMail());
    const refs = await new OutlookBridge().listContext();
    expect(refs).toHaveLength(1);
    expect(refs[0]).toMatchObject({
      id: 'AAMk-9',
      kind: 'mail-item',
      surface: 'outlook',
      live: true,
      title: 'Email: SLA concerns',
      preview: 'SLA concerns',
    });
  });

  it('falls back to a generic title/id and omits preview when there is no subject', async () => {
    active = installOutlook({ body: 'hi' });
    const refs = await new OutlookBridge().listContext();
    expect(refs[0]).toMatchObject({ id: 'outlook:item', title: 'Current email' });
    expect(refs[0]?.preview).toBeUndefined();
  });

  it('returns [] when there is no active mail item', async () => {
    active = installOutlook(undefined);
    expect(await new OutlookBridge().listContext()).toEqual([]);
  });

  it('returns [] when the mailbox is absent (wrong host)', async () => {
    active = installOutlook(readMail(), { noMailbox: true });
    expect(await new OutlookBridge().listContext()).toEqual([]);
  });
});

describe('OutlookBridge.resolveContext (host read → pure mapping)', () => {
  const ref: ContextRef = {
    id: 'AAMk-9',
    kind: 'mail-item',
    surface: 'outlook',
    title: 'Email: SLA concerns',
  };

  it('reads the active item and maps it to valid grounding context (HTML normalized)', async () => {
    active = installOutlook(readMail());
    const ctx = await new OutlookBridge().resolveContext(ref);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined).toContain('Subject: SLA concerns');
    expect(joined).toContain('pat@acme.com');
    expect(joined).toContain('raise the SLA');
    expect(joined).not.toContain('<strong>'); // tags stripped through the string path
  });

  it('returns [] when there is no active item to resolve', async () => {
    active = installOutlook(undefined);
    expect(await new OutlookBridge().resolveContext(ref)).toEqual([]);
  });

  it('rejects when the host body read fails (getAsync failure propagates)', async () => {
    active = installOutlook({ ...readMail(), bodyFails: true });
    await expect(new OutlookBridge().resolveContext(ref)).rejects.toThrow('body read failed');
  });
});

describe('OutlookBridge.revealContext', () => {
  it('opens the saved mail item in a host message form', async () => {
    active = installOutlook(readMail());
    const bridge = new OutlookBridge();
    const ref: ContextRef = {
      id: 'AAMk-9',
      kind: 'mail-item',
      surface: 'outlook',
      title: 'Email: SLA concerns',
    };

    expect(bridge.canRevealContext(ref)).toBe(true);
    await bridge.revealContext(ref);

    expect(active.openedMessages).toEqual(['AAMk-9']);
    expect(active.replyForms).toHaveLength(0);
    expect(active.newMessageForms).toHaveLength(0);
  });

  it('does not advertise an unsaved draft fallback ref as revealable', () => {
    active = installOutlook({ subject: 'Draft', body: 'wip' });
    expect(
      new OutlookBridge().canRevealContext({
        id: 'outlook:item',
        kind: 'mail-item',
        surface: 'outlook',
        title: 'Current email',
      }),
    ).toBe(false);
  });
});

describe('OutlookBridge.captureDocState (whole-item read)', () => {
  it('snapshots subject heading + from + body lines and bumps version each capture', async () => {
    active = installOutlook(readMail());
    const bridge = new OutlookBridge();
    const first = await bridge.captureDocState();
    expect(first).toBeDefined();
    if (!first) return;
    expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
    expect(first.surface).toBe('outlook');
    expect(first.version).toBe(1);
    expect(first.title).toBe('SLA concerns');

    const second = await bridge.captureDocState();
    expect(second?.version).toBe(2);
  });

  it('returns undefined when there is no active item', async () => {
    active = installOutlook(undefined);
    expect(await new OutlookBridge().captureDocState()).toBeUndefined();
  });

  it('returns undefined (no version bump) when the item yields no blocks', async () => {
    // No subject, no from, empty body → mailItemToDocStateBlocks returns [].
    active = installOutlook({ body: '   ' });
    const bridge = new OutlookBridge();
    expect(await bridge.captureDocState()).toBeUndefined();
    // Version did not increment on an empty capture: a subsequent real capture starts at 1.
    active.setItem(readMail());
    const snap = await bridge.captureDocState();
    expect(snap?.version).toBe(1);
  });
});

describe('OutlookBridge.searchDocument (scoped to the active item)', () => {
  it('returns matching body lines as valid context, scoped to the item', async () => {
    active = installOutlook(readMail());
    const ctx = await new OutlookBridge().searchDocument('raise the sla');
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const joined = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(joined.toLowerCase()).toContain('raise the sla');
  });

  it('returns [] for an empty query without reading the host', async () => {
    active = installOutlook({ ...readMail(), bodyFails: true });
    // A blank query short-circuits before any getAsync, so the failing body is never touched.
    expect(await new OutlookBridge().searchDocument('   ')).toEqual([]);
  });

  it('returns [] when there is no active item', async () => {
    active = installOutlook(undefined);
    expect(await new OutlookBridge().searchDocument('sla')).toEqual([]);
  });

  it('returns [] when the query matches nothing in the body', async () => {
    active = installOutlook(readMail());
    expect(await new OutlookBridge().searchDocument('nonexistent-token-xyz')).toEqual([]);
  });
});

describe('OutlookBridge.actuate reply-mail (reviewable reply form)', () => {
  it('opens a reply form with the planned html body and returns ok at reply-form', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(
      reply({ mail: { body: 'We can raise availability to 99.9%.' } }, 'chg-r'),
    );
    expect(res).toMatchObject({
      ok: true,
      changeId: asChangeId('chg-r'),
      kind: 'reply-mail',
      location: 'reply-form',
    });
    expect(active.replyForms).toEqual([{ htmlBody: 'We can raise availability to 99.9%.' }]);
  });

  it('rejects with no_item when there is no active mail to reply to (no form opened)', async () => {
    active = installOutlook(undefined);
    const res = await new OutlookBridge().actuate(reply({ mail: { body: 'hi' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_item' } });
    expect(active.replyForms).toHaveLength(0);
  });

  it('rejects with no_body when neither params.mail.body nor params.text is present', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(reply({ mail: { subject: 'Re: x' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_body' } });
    expect(active.replyForms).toHaveLength(0);
  });

  it('rejects with no_body when the body is only whitespace', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(reply({ text: '   ' }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_body' } });
    expect(active.replyForms).toHaveLength(0);
  });
});

describe('OutlookBridge.actuate create-mail (reviewable new draft)', () => {
  it('opens a new message form with subject/body and returns ok at new-message-form', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(
      compose({ mail: { subject: 'Follow-up on Q3', body: 'Summary below.' } }, 'chg-c'),
    );
    expect(res).toMatchObject({
      ok: true,
      changeId: asChangeId('chg-c'),
      kind: 'create-mail',
      location: 'new-message-form',
    });
    expect(active.newMessageForms).toEqual([
      { subject: 'Follow-up on Q3', htmlBody: 'Summary below.' },
    ]);
  });

  it('passes recipients through only when the agent supplied them', async () => {
    active = installOutlook(readMail());
    await new OutlookBridge().actuate(
      compose({ mail: { subject: 'Re: dates', to: ['pat@acme.com'] }, text: 'Confirming.' }),
    );
    expect(active.newMessageForms[0]).toEqual({
      subject: 'Re: dates',
      htmlBody: 'Confirming.',
      toRecipients: ['pat@acme.com'],
    });
  });

  it('rejects with no_mailbox when the host has no displayNewMessageForm', async () => {
    active = installOutlook(readMail(), { noNewMessageForm: true });
    const res = await new OutlookBridge().actuate(compose({ mail: { subject: 'x' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_mailbox' } });
    expect(active.newMessageForms).toHaveLength(0);
  });

  it('rejects with no_mailbox when there is no mailbox at all', async () => {
    active = installOutlook(readMail(), { noMailbox: true });
    const res = await new OutlookBridge().actuate(compose({ mail: { subject: 'x' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_mailbox' } });
  });

  it('rejects with no_subject when the subject is empty (no form opened)', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(compose({ mail: { body: 'orphan body' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'no_subject' } });
    expect(active.newMessageForms).toHaveLength(0);
  });
});

describe('OutlookBridge.actuate set-recipients (in-place draft edit)', () => {
  it('replaces To/Cc via setAsync and returns ok at recipients:set', async () => {
    active = installOutlook(draftMail({ to: ['old@acme.com'] }));
    const res = await new OutlookBridge().actuate(
      act('set-recipients', { mail: { to: ['a@acme.com'], cc: ['b@acme.com'] } }),
    );
    expect(res).toMatchObject({ ok: true, kind: 'set-recipients', location: 'recipients:set' });
    expect(active.recipientWrites).toEqual([
      { mode: 'set', field: 'to', addresses: ['a@acme.com'] },
      { mode: 'set', field: 'cc', addresses: ['b@acme.com'] },
    ]);
  });

  it('appends via addAsync when recipientMode=add and returns location recipients:add', async () => {
    active = installOutlook(draftMail({ to: ['old@acme.com'] }));
    const res = await new OutlookBridge().actuate(
      act('set-recipients', {
        mail: { to: ['new@acme.com'], recipientMode: 'add' },
      }),
    );
    expect(res).toMatchObject({ ok: true, location: 'recipients:add' });
    expect(active.recipientWrites).toEqual([
      { mode: 'add', field: 'to', addresses: ['new@acme.com'] },
    ]);
  });

  it('rejects with no_recipients when no usable address is supplied (nothing written)', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    for (const params of [{}, { mail: {} }, { mail: { to: [], cc: ['   '] } }]) {
      const res = await bridge.actuate(act('set-recipients', params));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_recipients' } });
    }
    expect(active.recipientWrites).toHaveLength(0);
  });

  it('validates every requested field up front so a missing field never partially applies', async () => {
    // A draft that exposes To but has NO cc field at all.
    active = installOutlook({ subject: 'd', body: 'x', draft: { to: [] } });
    const res = await new OutlookBridge().actuate(
      act('set-recipients', { mail: { to: ['a@acme.com'], cc: ['b@acme.com'] } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_compose' } });
    expect(active.recipientWrites).toHaveLength(0);
  });

  it('surfaces a host write failure as write_failed', async () => {
    active = installOutlook(draftMail({ recipientsFailsOn: 'bcc' }));
    const res = await new OutlookBridge().actuate(
      act('set-recipients', { mail: { bcc: ['x@acme.com'] } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'write_failed' } });
  });

  it('rejects with no_item when there is no active item at all', async () => {
    active = installOutlook(undefined);
    const res = await new OutlookBridge().actuate(
      act('set-recipients', { mail: { to: ['a@acme.com'] } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_item' } });
  });

  it('rejects with no_compose when the active item is a read-mode message', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(
      act('set-recipients', { mail: { to: ['a@acme.com'] } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_compose' } });
  });
});

describe('OutlookBridge.actuate add-attachment (in-place draft edit)', () => {
  it('attaches a base64 payload via addFileAttachmentFromBase64Async and records the minted id', async () => {
    active = installOutlook(draftMail());
    const base64 = 'aGVsbG8='; // "hello"
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { name: 'brief.txt', base64 } }, 'chg-att'),
    );
    expect(res).toMatchObject({
      ok: true,
      changeId: asChangeId('chg-att'),
      kind: 'add-attachment',
      location: 'draft-attachment:att-base64-1',
    });
    expect(active.attachmentsAdded).toEqual([
      { transport: 'base64', value: base64, name: 'brief.txt', isInline: false },
    ]);
  });

  it('derives the file name from an https uri and attaches via addFileAttachmentAsync', async () => {
    active = installOutlook(draftMail());
    const res = await new OutlookBridge().actuate(
      act('add-attachment', {
        attachment: { uri: 'https://contoso.example/reports/q3.pdf?token=x' },
      }),
    );
    expect(res).toMatchObject({ ok: true, location: 'draft-attachment:att-uri-1' });
    expect(active.attachmentsAdded).toEqual([
      {
        transport: 'uri',
        value: 'https://contoso.example/reports/q3.pdf?token=x',
        name: 'q3.pdf',
        isInline: false,
      },
    ]);
  });

  it('attaches another mail item via addItemAttachmentAsync when itemId is given with a name', async () => {
    active = installOutlook(draftMail());
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { itemId: 'AAMk-src', name: 'Forwarded.msg' } }),
    );
    expect(res).toMatchObject({ ok: true, location: 'draft-attachment:att-item-1' });
    expect(active.attachmentsAdded).toEqual([
      { transport: 'item', value: 'AAMk-src', name: 'Forwarded.msg', isInline: false },
    ]);
  });

  it('rejects with no_attachment when the request carries no usable payload or name', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    for (const params of [
      {},
      { attachment: {} },
      { attachment: { name: 'blank.txt', base64: '   ' } },
      { attachment: { itemId: 'AAMk-src' } },
    ]) {
      const res = await bridge.actuate(act('add-attachment', params));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_attachment' } });
    }
    expect(active.attachmentsAdded).toHaveLength(0);
  });

  it('rejects an oversized base64 payload before any host call and never echoes it', async () => {
    active = installOutlook(draftMail());
    const oversized = `${'Q'.repeat(1024)}${'Q'.repeat(MAX_ATTACHMENT_BASE64_CHARS)}`;
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { name: 'big.bin', base64: oversized } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'attachment_too_large' } });
    expect(JSON.stringify(res.error)).not.toContain('QQQQ');
    expect(active.attachmentsAdded).toHaveLength(0);
  });

  it('rejects invalid base64 (bad charset/padding) as invalid_attachment', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    for (const base64 of ['abc', 'not+base64!', 'aGVs*b64=']) {
      const res = await bridge.actuate(
        act('add-attachment', { attachment: { name: 'x.bin', base64 } }),
      );
      expect(res).toMatchObject({ ok: false, error: { code: 'invalid_attachment' } });
    }
    expect(active.attachmentsAdded).toHaveLength(0);
  });

  it('rejects a non-https uri scheme (local-file exfil guard) as invalid_attachment', async () => {
    active = installOutlook(draftMail());
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { uri: 'file:///etc/hosts' } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'invalid_attachment' } });
    expect(active.attachmentsAdded).toHaveLength(0);
  });

  it('surfaces a host upload failure as write_failed', async () => {
    active = installOutlook(draftMail({ attachmentAddFails: true }));
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { name: 'brief.txt', base64: 'aGVsbG8=' } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'write_failed' } });
  });

  it('rejects with no_compose on a read-mode item', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate(
      act('add-attachment', { attachment: { name: 'brief.txt', base64: 'aGVsbG8=' } }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_compose' } });
  });
});

describe('OutlookBridge.actuate set-body (in-place draft edit)', () => {
  it('writes HTML by default via Body.setAsync and returns ok at draft-body', async () => {
    active = installOutlook(draftMail());
    const res = await new OutlookBridge().actuate(
      act('set-body', { mail: { body: '<p>Raised to 99.9%.</p>' } }),
    );
    expect(res).toMatchObject({ ok: true, kind: 'set-body', location: 'draft-body' });
    expect(active.bodySets).toEqual([{ data: '<p>Raised to 99.9%.</p>', coercionType: 'html' }]);
  });

  it('falls back to params.text with Text coercion, overridable by mail.coercion', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    await bridge.actuate(act('set-body', { text: 'Plain prose.' }));
    expect(active.bodySets[0]).toEqual({ data: 'Plain prose.', coercionType: 'text' });

    await bridge.actuate(act('set-body', { html: '<p>Rich</p>' }));
    expect(active.bodySets[1]).toEqual({ data: '<p>Rich</p>', coercionType: 'html' });

    await bridge.actuate(act('set-body', { mail: { body: '<p>x</p>', coercion: 'text' } }));
    expect(active.bodySets[2]).toEqual({ data: '<p>x</p>', coercionType: 'text' });
  });

  it('rejects with no_body when no usable body is supplied (nothing written)', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    for (const params of [{}, { mail: {} }, { text: '   ' }]) {
      const res = await bridge.actuate(act('set-body', params));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_body' } });
    }
    expect(active.bodySets).toHaveLength(0);
  });

  it('surfaces a host write failure as write_failed', async () => {
    active = installOutlook(draftMail({ bodySetFails: true }));
    const res = await new OutlookBridge().actuate(act('set-body', { mail: { body: '<p>x</p>' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'write_failed' } });
  });
});

describe('OutlookBridge.actuate set-subject (in-place draft edit)', () => {
  it('replaces the subject via Subject.setAsync and returns ok at draft-subject', async () => {
    active = installOutlook(draftMail());
    const res = await new OutlookBridge().actuate(
      act('set-subject', { mail: { subject: 'Updated project plan' } }),
    );
    expect(res).toMatchObject({
      ok: true,
      kind: 'set-subject',
      location: 'draft-subject',
    });
    expect(active.subjectSets).toEqual([{ subject: 'Updated project plan' }]);
  });

  it('rejects with no_subject when the subject is missing or blank (nothing written)', async () => {
    active = installOutlook(draftMail());
    const bridge = new OutlookBridge();
    for (const params of [{}, { mail: {} }, { mail: { subject: '   ' } }]) {
      const res = await bridge.actuate(act('set-subject', params));
      expect(res).toMatchObject({ ok: false, error: { code: 'no_subject' } });
    }
    expect(active.subjectSets).toHaveLength(0);
  });

  it('surfaces a host write failure as write_failed', async () => {
    active = installOutlook(draftMail({ subjectSetFails: true }));
    const res = await new OutlookBridge().actuate(act('set-subject', { mail: { subject: 'x' } }));
    expect(res).toMatchObject({ ok: false, error: { code: 'write_failed' } });
  });
});

describe('OutlookBridge.actuate unsupported', () => {
  it('returns an unsupported error for a kind Outlook cannot do', async () => {
    active = installOutlook(readMail());
    const res = await new OutlookBridge().actuate({
      changeId: asChangeId('chg-u'),
      kind: 'tracked-change',
      surface: 'outlook',
      params: { text: 'x' },
    });
    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.kind).toBe('tracked-change');
  });
});

describe('OutlookBridge.watch (ItemChanged event wiring)', () => {
  it('emits mail-received for a saved item and mail-compose for an unsaved draft', async () => {
    active = installOutlook(readMail());
    const events: HostEvent[] = [];
    const bridge = new OutlookBridge();
    const unsub = bridge.watch((e) => events.push(e));

    // Saved item active → mail-received with its id.
    active.fireItemChanged();
    // Switch to an unsaved draft (no itemId) → mail-compose.
    active.setItem({ subject: 'Draft', body: 'wip' });
    active.fireItemChanged();

    expect(events).toEqual([{ type: 'mail-received', id: 'AAMk-9' }, { type: 'mail-compose' }]);

    // Unsubscribe detaches the handler so no further events fire.
    unsub();
    active.fireItemChanged();
    expect(events).toHaveLength(2);
  });

  it('emits nothing when ItemChanged fires with no active item', async () => {
    active = installOutlook(readMail());
    const events: HostEvent[] = [];
    new OutlookBridge().watch((e) => events.push(e));
    active.setItem(undefined);
    active.fireItemChanged();
    expect(events).toHaveLength(0);
  });

  it('returns a no-op unsubscribe and registers nothing when the mailbox is absent', async () => {
    active = installOutlook(readMail(), { noMailbox: true });
    const events: HostEvent[] = [];
    const unsub = new OutlookBridge().watch((e) => events.push(e));
    expect(active.handlers).toHaveLength(0);
    expect(() => unsub()).not.toThrow();
  });

  it('returns a no-op unsubscribe when addHandlerAsync throws (registration failure)', async () => {
    active = installOutlook(readMail(), { throwOnAdd: true });
    const unsub = new OutlookBridge().watch(() => {});
    expect(active.handlers).toHaveLength(0);
    expect(() => unsub()).not.toThrow();
  });

  it('swallows a handler error so a bad event-source notification never breaks the host', async () => {
    active = installOutlook(readMail());
    const bridge = new OutlookBridge();
    // The emit callback throws; the bridge's internal try/catch must isolate it.
    const unsub = bridge.watch(() => {
      throw new Error('downstream emit blew up');
    });
    expect(() => active!.fireItemChanged()).not.toThrow();
    unsub();
  });

  it('swallows a teardown error from removeHandlerAsync (best-effort unsubscribe)', async () => {
    active = installOutlook(readMail(), { throwOnRemove: true });
    const unsub = new OutlookBridge().watch(() => {});
    expect(active.handlers).toHaveLength(1);
    expect(() => unsub()).not.toThrow();
  });
});
