#!/usr/bin/env bun
/**
 * Design check: the task pane follows packages/web-shell/DESIGN.md.
 *
 *   bun scripts/check-design.ts            fail when any file has more violations than its baseline
 *   bun scripts/check-design.ts --update   accept the current counts (after a clean-up lowers them)
 *
 * Rules (CSS in packages/web-shell/src/taskpane, components in .../components):
 *   font-size     a size that is not a --fz-* token (or inherit)
 *   radius        a border-radius that is not var(--r*), 0, 50% or 9999px (round knobs)
 *   color         a hex / rgb() / hsl() colour outside a custom-property declaration
 *   button        a <button> that is not one of the recipes in DESIGN.md § Buttons (.btn, .link, …)
 *
 * Old code is held at the counts in packages/web-shell/design-baseline.json, so lint stays green
 * today and only new drift fails. Lower a count by cleaning up, then run --update.
 */
import { readFileSync, readdirSync, writeFileSync, existsSync } from 'node:fs';
import { join, relative } from 'node:path';

const ROOT = join(import.meta.dir, '..');
const PANE = join(ROOT, 'packages', 'web-shell', 'src', 'taskpane');
const BASELINE = join(ROOT, 'packages', 'web-shell', 'design-baseline.json');

type Rule = 'font-size' | 'radius' | 'color' | 'button';
type Counts = Record<string, Partial<Record<Rule, number>>>;
interface Hit {
  file: string;
  line: number;
  rule: Rule;
  text: string;
}

const FONT_OK = /^(var\(--fz-[a-z]+\)|inherit)$/;
const RADIUS_OK = /^((var\(--r[a-z-]*\)|0|50%|9999px)\s*)+$/;
const COLOR = /#[0-9a-fA-F]{3,8}\b|\brgba?\(|\bhsla?\(/;
// The button recipes in DESIGN.md § Buttons. A new kind of button is a new recipe there first.
const RECIPES =
  /(^|[\s'"`])(btn|link|tw-rail-item|tw-rail-toggle|tw-head-btn|tw-modal-close|cite|suggestion|workspace-action|palette-item)([\s'"`]|$)/;

function cssHits(file: string): Hit[] {
  const hits: Hit[] = [];
  readFileSync(file, 'utf8')
    .split('\n')
    .forEach((raw, i) => {
      const line = raw.replace(/\/\*.*?\*\//g, '').trim();
      if (!line || line.startsWith('--') || line.startsWith('/*') || line.startsWith('*')) return;
      const decl = /^([a-z-]+)\s*:\s*(.+?);?$/.exec(line);
      if (!decl) return;
      const [, prop, value] = decl as unknown as [string, string, string];
      const v = value.replace(/\s*!important$/, '').trim();
      const hit = (rule: Rule): void => {
        hits.push({ file, line: i + 1, rule, text: line });
      };
      if (prop === 'font-size' && !FONT_OK.test(v)) hit('font-size');
      if (prop === 'border-radius' && !RADIUS_OK.test(v)) hit('radius');
      if (COLOR.test(v)) hit('color');
    });
  return hits;
}

function tsxHits(file: string): Hit[] {
  const src = readFileSync(file, 'utf8');
  const hits: Hit[] = [];
  for (const m of src.matchAll(/<button\b([^>]*?)>/gs)) {
    const cls = /className=(?:"([^"]*)"|\{`([^`]*)`\}|\{([^}]*)\})/.exec(m[1] ?? '');
    const value = cls ? (cls[1] ?? cls[2] ?? cls[3] ?? '') : '';
    if (RECIPES.test(value)) continue;
    const line = src.slice(0, m.index).split('\n').length;
    hits.push({ file, line, rule: 'button', text: m[0].replace(/\s+/g, ' ').slice(0, 90) });
  }
  return hits;
}

const files = [
  ...readdirSync(PANE)
    .filter((f) => f.endsWith('.css'))
    .map((f) => join(PANE, f)),
  ...readdirSync(join(PANE, 'components'))
    .filter((f) => f.endsWith('.tsx'))
    .map((f) => join(PANE, 'components', f)),
];
const hits = files.flatMap((f) => (f.endsWith('.css') ? cssHits(f) : tsxHits(f)));

const counts: Counts = {};
for (const h of hits) {
  const key = relative(ROOT, h.file);
  const entry = (counts[key] ??= {});
  entry[h.rule] = (entry[h.rule] ?? 0) + 1;
}

if (process.argv.includes('--update')) {
  writeFileSync(BASELINE, `${JSON.stringify(counts, null, 2)}\n`);
  console.log(`design baseline updated: ${hits.length} known violations`);
  process.exit(0);
}

const baseline: Counts = existsSync(BASELINE) ? JSON.parse(readFileSync(BASELINE, 'utf8')) : {};
let failed = false;
for (const [file, rules] of Object.entries(counts)) {
  for (const [rule, n] of Object.entries(rules) as [Rule, number][]) {
    const allowed = baseline[file]?.[rule] ?? 0;
    if (n <= allowed) continue;
    failed = true;
    console.error(`\n${file}: ${n} ${rule} violations (baseline ${allowed})`);
    for (const h of hits.filter((x) => relative(ROOT, x.file) === file && x.rule === rule)) {
      console.error(`  ${h.line}: ${h.text}`);
    }
  }
}
if (failed) {
  console.error(
    '\nDesign check failed. Use the tokens and the .btn / .field recipes in packages/web-shell/DESIGN.md.',
  );
  process.exit(1);
}
const known = Object.values(baseline)
  .flatMap((r) => Object.values(r))
  .reduce((a, b) => a + (b ?? 0), 0);
console.log(`design check passed (${known} known violations held at baseline)`);
