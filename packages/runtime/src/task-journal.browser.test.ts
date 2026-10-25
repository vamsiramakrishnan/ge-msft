import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest';
import { chromium, type Browser, type BrowserContext } from 'playwright';
import { createServer, type Server } from 'node:http';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import type { ActuationRequest, ActuationResult } from '@ge/contracts';

interface BrowserState {
  started: boolean;
  release(): void;
  pending: Promise<ActuationResult>;
}
describe.skipIf(process.env.GE_BROWSER_CONTRACTS !== '1')('IndexedDB cross-pane contract', () => {
  let browser: Browser;
  let context: BrowserContext;
  let server: Server;
  let directory: string;
  let origin: string;
  const scope = 'sha256:' + 'a'.repeat(64);
  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'ge-journal-browser-'));
    const bundle = join(directory, 'journal.js');
    execFileSync(
      'bun',
      [
        'build',
        fileURLToPath(new URL('./task-journal.ts', import.meta.url)),
        '--target',
        'browser',
        '--outfile',
        bundle,
      ],
      { stdio: 'pipe' },
    );
    const source = await readFile(bundle);
    server = createServer((request, response) => {
      response.setHeader(
        'Content-Type',
        request.url === '/journal.js' ? 'text/javascript' : 'text/html',
      );
      response.end(
        request.url === '/journal.js' ? source : '<!doctype html><title>Journal contract</title>',
      );
    });
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
    const address = server.address();
    if (!address || typeof address === 'string') throw new Error('Local test origin unavailable.');
    origin = `http://127.0.0.1:${address.port}`;
    browser = await chromium.launch({ headless: true });
  }, 30000);
  beforeEach(async () => {
    context = await browser.newContext();
  });
  afterEach(async () => {
    await context?.close();
  });
  afterAll(async () => {
    await browser?.close();
    if (server?.listening)
      await new Promise<void>((resolve, reject) =>
        server.close((error) => (error ? reject(error) : resolve())),
      );
    if (directory) await rm(directory, { recursive: true, force: true });
  });
  it('serializes claims across panes and retains the receipt after reopening', async () => {
    const first = await context.newPage();
    const second = await context.newPage();
    await Promise.all([first.goto(origin), second.goto(origin)]);
    await first.evaluate(async (scope) => {
      const url = '/journal.js';
      const module = (await import(url)) as typeof import('./task-journal.js');
      const journal = new module.TaskJournal(new module.IndexedDbTaskJournalStore(), scope);
      const request: ActuationRequest = {
        changeId: 'first' as ActuationRequest['changeId'],
        surface: 'word',
        kind: 'tracked-change',
        params: { target: { matchText: 'before' }, text: 'after' },
      };
      const state: BrowserState = {
        started: false,
        release: () => undefined,
        pending: Promise.resolve({ ok: false, changeId: request.changeId, kind: request.kind }),
      };
      (window as unknown as { journalTest: BrowserState }).journalTest = state;
      state.pending = journal.execute(
        'task-1',
        request,
        () =>
          new Promise<ActuationResult>((resolve) => {
            state.started = true;
            state.release = () =>
              resolve({ ok: true, changeId: request.changeId, kind: request.kind });
          }),
      );
    }, scope);
    await first.waitForFunction(
      () => (window as unknown as { journalTest: BrowserState }).journalTest.started,
    );
    const conflict = await second.evaluate(async (scope) => {
      const url = '/journal.js';
      const module = (await import(url)) as typeof import('./task-journal.js');
      const journal = new module.TaskJournal(new module.IndexedDbTaskJournalStore(), scope);
      let dispatched = false;
      try {
        await journal.execute(
          'task-2',
          {
            changeId: 'second' as ActuationRequest['changeId'],
            surface: 'word',
            kind: 'tracked-change',
            params: { text: 'duplicate' },
          },
          async () => {
            dispatched = true;
            throw new Error('must not dispatch');
          },
        );
      } catch (error) {
        return { dispatched, error: String(error) };
      }
      return { dispatched, error: '' };
    }, scope);
    expect(conflict).toMatchObject({
      dispatched: false,
      error: expect.stringContaining('unresolved'),
    });
    await first.evaluate(async () => {
      const state = (window as unknown as { journalTest: BrowserState }).journalTest;
      state.release();
      await state.pending;
    });
    await first.close();
    const reopened = await context.newPage();
    await reopened.goto(origin);
    const history = await reopened.evaluate(async (scope) => {
      const url = '/journal.js';
      const module = (await import(url)) as typeof import('./task-journal.js');
      const journal = new module.TaskJournal(new module.IndexedDbTaskJournalStore(), scope);
      return { durable: journal.durable, entries: await journal.inspect() };
    }, scope);
    expect(history).toMatchObject({ durable: true, entries: [{ id: 'first', state: 'applied' }] });
  }, 30000);
  it('aborts invalid updates atomically', async () => {
    const page = await context.newPage();
    await page.goto(origin);
    const result = await page.evaluate(async (scope) => {
      const url = '/journal.js';
      const module = (await import(url)) as typeof import('./task-journal.js');
      const store = new module.IndexedDbTaskJournalStore();
      await store.update(scope, () => []);
      let rejected = false;
      try {
        await store.update(scope, () => [{}] as never);
      } catch {
        rejected = true;
      }
      return { rejected, history: await store.read(scope) };
    }, scope);
    expect(result).toEqual({ rejected: true, history: [] });
  }, 30000);
});
