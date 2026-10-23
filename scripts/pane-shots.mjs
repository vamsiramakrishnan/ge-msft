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

const browser = await chromium.launch({ channel: 'chrome' }).catch(() => chromium.launch());
const page = await browser.newPage({ viewport: { width: 1400, height: 1000 }, deviceScaleFactor: 2 });
for (const width of [320, 360, 480]) {
  const query = [`surface=${surface}`, 'demo=1', args.query].filter(Boolean).join('&');
  await page.goto(`${base}/?${query}`);
  await page.selectOption('select[aria-label="Pane width"]', `${width}`);
  await page.waitForTimeout(400);
  if (args.panel) {
    await page.click(`.tw-rail-item[data-panel="${args.panel}"]`);
    await page.waitForTimeout(200);
  }
  const file = join(out, `${surface}${args.panel ? `-${args.panel}` : ''}-${width}.png`);
  await page.locator('.preview-frame').screenshot({ path: file });
  console.log(file);
}
await browser.close();
