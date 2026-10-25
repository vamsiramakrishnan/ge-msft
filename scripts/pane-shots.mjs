#!/usr/bin/env node
/**
 * Screenshot the task-pane preview at the three standard widths (DESIGN.md § 7).
 *
 *   node scripts/pane-shots.mjs [--url http://localhost:17100] [--surface excel] [--panel settings]
 *                               [--query resume=1] [--out <dir>]
 *
 * Needs the preview server running (`bun run --filter @ge/web-shell preview`). Writes
 * <out>/<surface>[-<panel>]-<width>.png (default out: $TMPDIR/pane-shots) and prints each path.
 * --panel opens a sidebar panel first: actions | skills | sessions | settings.
 */
import { mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { chromium } from 'playwright';

const args = {};
const argv = process.argv.slice(2);
for (let i = 0; i < argv.length; i += 2) {
  if (!argv[i]?.startsWith('--')) throw new Error(`expected a --flag, got ${argv[i]}`);
  args[argv[i].slice(2)] = argv[i + 1];
}
const base = args.url ?? 'http://localhost:17100';
const surface = args.surface ?? 'excel';
const out = args.out ?? join(tmpdir(), 'pane-shots');
mkdirSync(out, { recursive: true });

const browser = process.env.GE_BROWSER_EXECUTABLE
  ? await chromium.launch({ executablePath: process.env.GE_BROWSER_EXECUTABLE })
  : await chromium.launch({ channel: 'chrome' }).catch(() => chromium.launch());
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, deviceScaleFactor: 2 });
for (const width of [320, 360, 480]) {
  const query = [`surface=${surface}`, args.demo === '0' ? '' : 'demo=1', args.query].filter(Boolean).join('&');
  await page.goto(`${base}/?${query}`);
  await page.selectOption('select[aria-label="Pane width"]', `${width}`);
  await page.waitForTimeout(400);
  if (args.panel) {
    await page.click(`.tw-rail-item[data-panel="${args.panel}"]`);
    await page.waitForTimeout(200);
  }
  if (args.expand) {
    const disclosure = page.locator('.office-workbench').filter({ has: page.locator('summary', { hasText: args.expand }) }).first();
    await disclosure.locator(':scope > summary').click();
    if (args.fill === '1') {
      for (const field of await disclosure.locator('textarea:not([readonly])').all()) await field.fill('Sample instructions for this workflow.');
      const output = disclosure.locator('select');
      if (await output.count()) await output.selectOption({ index: 1 });
    }
    await disclosure.scrollIntoViewIfNeeded();
  }
  if (args.scroll === 'interrupted') {
    if (args.check === '1') await page.getByRole('checkbox', { name: 'I stopped other running panes' }).check();
    await page.getByRole('heading', { name: 'Inspect interrupted changes' }).scrollIntoViewIfNeeded();
  }
  if (args.scroll === 'report') await page.getByRole('button', {name: 'Copy native report', exact: true}).scrollIntoViewIfNeeded();
  if (args.commands === '1') {
    const exact = page.locator('.saved-office-task-inspection details');
    await exact.locator('summary').click();
    await exact.scrollIntoViewIfNeeded();
  }
  if (args.create === '1') {
    const create = page.locator('.saved-office-task-create');
    await create.locator('summary').click();
    await create.getByLabel('Task ID', {exact:true}).fill('month-close-reviewed-memo');
    await create.getByLabel('Exact commands', {exact:true}).fill('/fill-content-control contentControlId="42" text="Reviewed close summary. Source amounts retain their supplied precision."\nfinish when=verified');
    if(args.attest==='1')await create.getByRole('checkbox').check();
    await create.scrollIntoViewIfNeeded();
  }
  if (args.approval === '1') {
    await page.locator('.plan-effect-head').first().click();
    await page.locator('.plan-effect-detail').first().scrollIntoViewIfNeeded();
  }
  if(args.tag&&!/^[a-z0-9-]{1,64}$/.test(args.tag))throw new Error('Use a short public screenshot tag.');
  const file = join(out, `${surface}${args.panel ? `-${args.panel}` : ''}${args.expand ? `-${args.expand.replaceAll(' ', '-')}` : ''}${args.scroll ? `-${args.scroll}` : ''}${args.tag?`-${args.tag}`:''}-${width}.png`);
  await page.locator('.preview-frame').screenshot({ path: file });
  console.log(file);
}
await browser.close();
