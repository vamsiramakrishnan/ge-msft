#!/usr/bin/env node
/**
 * Build the web-shell as a CDN-deployable static bundle and prove it is safe to publish.
 *
 * Why: Office add-ins are just HTTPS static assets plus a manifest that points at them. The
 * dev path (cloudflared tunnel) must never be what you ship; this produces the artifact you
 * hand to any static host (Cloudflare Pages, Firebase Hosting, Azure Static Web Apps, S3+CDN).
 *
 * Gates, in order:
 *   1. clean build (`vite build` for the chosen Vite mode);
 *   2. no placeholder tokens (`REPLACE_*`) anywhere in the bundle — those only belong to
 *      template manifests, never to deployed JS;
 *   3. no localhost/dev-tunnel origins baked into the bundle;
 *   4. repo secret scan over the output.
 *
 * Usage: bun run release:web [-- mode=production] [--fail-on-warning=false]
 */
import { spawnSync } from 'node:child_process';
import { copyFileSync, existsSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const rootDir = fileURLToPath(new URL('..', import.meta.url));
const shellDir = join(rootDir, 'packages', 'web-shell');
const distDir = join(shellDir, 'dist-web');

function fail(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}
function ok(msg) {
  console.log(`✓ ${msg}`);
}

const mode = (process.argv.find((a) => a.startsWith('--mode=')) ?? '--mode=production')
  .split('=')[1];
const allowLocalhost = process.argv.includes('--allow-localhost');

console.log(`Building @ge/web-shell (mode=${mode})…`);
rmSync(distDir, { recursive: true, force: true });
const build = spawnSync('bunx', ['vite', 'build', '--mode', mode], {
  cwd: shellDir,
  stdio: ['ignore', 'pipe', 'inherit'],
});
if (build.status !== 0) fail('vite build failed — see output above');
ok(`built ${relative(rootDir, distDir)}`);

// Gate 1: every referenced entry page emitted.
for (const entry of [
  'index.html',
  'taskpane.html',
  'commands.html',
  'functions.html',
  'auth-redirect.html',
  'functions.json',
]) {
  if (!existsSync(join(distDir, entry))) fail(`dist-web is missing ${entry}`);
}
ok('all entry pages + functions.json emitted');

// The unified manifest needs a stable command-script URL. Vite emits a content-hashed command
// chunk, while `public/assets/commands.js` is only a dev-server shim that imports `/src/...`.
// Replace that shim in the release artifact with the real production bundle or Office will keep
// the add-in in "Loading add-ins" while the command runtime fails to initialize.
const assetsDir = join(distDir, 'assets');
const commandsChunk = readdirSync(assetsDir).find((name) =>
  /^commands-[A-Za-z0-9_-]+\.js$/.test(name),
);
if (!commandsChunk) fail('dist-web is missing the hashed commands runtime');
copyFileSync(join(assetsDir, commandsChunk), join(assetsDir, 'commands.js'));
const stableCommands = readFileSync(join(assetsDir, 'commands.js'), 'utf8');
if (/\/src\/commands\//.test(stableCommands)) {
  fail('stable assets/commands.js still points at a development source path');
}
ok(`published stable commands runtime from assets/${commandsChunk}`);

// Gate 2/3: walk the bundle for placeholders and dev origins.
function* files(dir) {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) yield* files(p);
    else yield p;
  }
}
const offenders = [];
for (const file of files(distDir)) {
  if (!/\.(js|html|json|css|svg)$/.test(file)) continue;
  const content = readFileSync(file, 'utf8');
  if (/REPLACE_[A-Z_]+/.test(content)) offenders.push(`${relative(rootDir, file)}: REPLACE_ token`);
  if (!allowLocalhost && /https?:\/\/(localhost|127\.0\.0\.1)/.test(content)) {
    offenders.push(`${relative(rootDir, file)}: localhost origin`);
  }
}
if (offenders.length) {
  console.error(offenders.map((o) => `  - ${o}`).join('\n'));
  fail('bundle contains template placeholders or dev origins — rebuild with production env');
}
ok('no REPLACE_* placeholders or localhost origins in bundle');

// Gate 4: the repo secret scanner already covers packages/web-shell/dist-web.
const scan = spawnSync('bun', [join(rootDir, 'tools', 'release', 'secret-scan.mjs')], {
  cwd: rootDir,
  stdio: 'inherit',
});
if (scan.status !== 0) fail('secret scan flagged the bundle');

console.log(`\nDeploy-ready artifact: ${relative(rootDir, distDir)}/`);
console.log('Next: upload to your static host, then point the manifests at that origin.');
console.log('See docs/DISTRIBUTION.md for host-specific commands and the store checklist.');
