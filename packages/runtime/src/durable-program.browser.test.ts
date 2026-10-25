import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser } from 'playwright';
import { createServer, type Server } from 'node:http';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import type * as DurableModule from './durable-program.js';

describe.skipIf(process.env.GE_BROWSER_CONTRACTS !== '1')(
  'durable program browser checkpoint contract',
  () => {
    let browser: Browser;
    let server: Server;
    let directory: string;
    let origin: string;
    beforeAll(async () => {
      directory = await mkdtemp(join(tmpdir(), 'ge-program-browser-'));
      const bundle = join(directory, 'program.js');
      execFileSync(
        'bun',
        [
          'build',
          fileURLToPath(new URL('./durable-program.ts', import.meta.url)),
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
          request.url === '/program.js' ? 'text/javascript' : 'text/html',
        );
        response.end(
          request.url === '/program.js'
            ? source
            : '<!doctype html><title>Durable contract</title><script type="module">import * as durable from "/program.js";window.GEDurable=durable;</script>',
        );
      });
      await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
      const address = server.address();
      if (!address || typeof address === 'string') throw new Error('Test origin missing.');
      origin = `http://127.0.0.1:${address.port}`;
      browser = await chromium.launch({
        headless: true,
        executablePath: process.env.GE_BROWSER_EXECUTABLE,
      });
    }, 30000);
    afterAll(async () => {
      await browser?.close();
      if (server?.listening)
        await new Promise<void>((resolve, reject) =>
          server.close((error) => (error ? reject(error) : resolve())),
        );
      if (directory) await rm(directory, { recursive: true, force: true });
    });
    it('serializes stage claims across real panes and retains a verified checkpoint after reopening', async () => {
      const context = await browser.newContext();
      try {
        const first = await context.newPage();
        const second = await context.newPage();
        await Promise.all([first.goto(origin), second.goto(origin)]);
        const scope = 'sha256:' + 'a'.repeat(64);
        await first.evaluate(async (scope) => {
          const module = (globalThis as unknown as { GEDurable: typeof DurableModule }).GEDurable;
          const runner = new module.DurableProgramRunner(
            new module.IndexedDbDurableProgramStore(),
            scope,
          );
          let release!: () => void;
          const barrier = new Promise<void>((resolve) => {
            release = resolve;
          });
          const pending = (async () => {
            for await (const _event of runner.run(
              {
                version: 1,
                id: 'browser',
                surface: 'word',
                stages: [{ id: 'memo', dependsOn: [], program: 'suggest "old" => "new"' }],
              },
              async function* (_stage, checkpointEffect) {
                yield* [];
                await checkpointEffect({
                  changeId: 'native-change' as never,
                  kind: 'tracked-change',
                  surface: 'word',
                  params: { target: { matchText: 'old' }, text: 'new' },
                });
                (globalThis as unknown as { stageStarted: boolean }).stageStarted = true;
                await barrier;
                return [
                  {
                    ok: true,
                    kind: 'tracked-change' as const,
                    changeId: 'native-change' as never,
                    verification: { status: 'verified' as const },
                  },
                ];
              },
            )) {
              /* consume */
            }
          })();
          (
            globalThis as unknown as { stageRelease: () => void; stagePending: Promise<void> }
          ).stageRelease = release;
          (globalThis as unknown as { stagePending: Promise<void> }).stagePending = pending;
        }, scope);
        await first.waitForFunction(
          () => (globalThis as unknown as { stageStarted: boolean }).stageStarted === true,
        );
        expect(
          await second.evaluate(async (scope) => {
            const module = (globalThis as unknown as { GEDurable: typeof DurableModule }).GEDurable;
            const runner = new module.DurableProgramRunner(
              new module.IndexedDbDurableProgramStore(),
              scope,
            );
            let executed = false;
            try {
              for await (const _event of runner.run(
                {
                  version: 1,
                  id: 'browser',
                  surface: 'word',
                  stages: [{ id: 'memo', dependsOn: [], program: 'suggest "old" => "new"' }],
                },
                async function* () {
                  yield* [];
                  executed = true;
                  return [];
                },
              )) {
                /* consume */
              }
              return 'unexpected';
            } catch {
              return executed ? 'dispatched' : 'blocked';
            }
          }, scope),
        ).toBe('blocked');
        await first.evaluate(async () => {
          const state = globalThis as unknown as {
            stageRelease: () => void;
            stagePending: Promise<void>;
          };
          state.stageRelease();
          await state.stagePending;
        });
        await first.close();
        await second.reload();
        expect(
          await second.evaluate(async (scope) => {
            const module = (globalThis as unknown as { GEDurable: typeof DurableModule }).GEDurable;
            const runner = new module.DurableProgramRunner(
              new module.IndexedDbDurableProgramStore(),
              scope,
            );
            return (await runner.inspect('browser'))?.stages[0];
          }, scope),
        ).toMatchObject({
          state: 'verified',
          receipt: { changeId: 'native-change', verification: { status: 'verified' } },
        });
      } finally {
        await context.close();
      }
    });
  },
);
