#!/usr/bin/env bun
import { mkdir, writeFile, stat, access, readdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawn } from 'node:child_process';
import { chromium } from 'playwright';
import { GIFEncoder, quantize, applyPalette } from 'gifenc';
import { PNG } from 'pngjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_OUT = path.join(ROOT, 'docs/assets/readme/sidepanes');
const DEFAULT_PORT = 17101;
const SURFACES = ['word', 'excel', 'powerpoint', 'outlook', 'onenote', 'teams'];

const args = parseArgs(process.argv.slice(2));

if (args.help) {
  printHelp();
  process.exit(0);
}

const selectedSurfaces = args.surface
  ? args.surface
      .split(',')
      .map((s) => s.trim().toLowerCase())
      .filter(Boolean)
  : SURFACES;

for (const surface of selectedSurfaces) {
  if (!SURFACES.includes(surface)) {
    throw new Error(`Unknown surface "${surface}". Expected one of: ${SURFACES.join(', ')}`);
  }
}

const outDir = path.resolve(ROOT, args.out ?? DEFAULT_OUT);
const port = Number(args.port ?? DEFAULT_PORT);
const baseUrl = `http://127.0.0.1:${port}`;

if (args.check) {
  await checkExisting(outDir, selectedSurfaces);
  process.exit(0);
}

await mkdir(outDir, { recursive: true });

let server;
let browser;

try {
  server = startPreviewServer(port);
  await waitForPreview(baseUrl);

  browser = await chromium.launch(await chromiumLaunchOptions());
  for (const surface of selectedSurfaces) {
    const file = path.join(outDir, `${surface}.gif`);
    await renderSurfaceGif(browser, baseUrl, surface, file);
    const size = await stat(file);
    console.log(`rendered ${path.relative(ROOT, file)} (${Math.round(size.size / 1024)} KiB)`);
  }

  await writeManifest(outDir, selectedSurfaces);
} finally {
  if (browser) await browser.close();
  if (server) await stopServer(server);
}

function parseArgs(argv) {
  const parsed = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--help' || arg === '-h') parsed.help = true;
    else if (arg === '--check') parsed.check = true;
    else if (arg === '--surface') parsed.surface = argv[++i];
    else if (arg.startsWith('--surface=')) parsed.surface = arg.slice('--surface='.length);
    else if (arg === '--out') parsed.out = argv[++i];
    else if (arg.startsWith('--out=')) parsed.out = arg.slice('--out='.length);
    else if (arg === '--port') parsed.port = argv[++i];
    else if (arg.startsWith('--port=')) parsed.port = arg.slice('--port='.length);
    else throw new Error(`Unknown argument: ${arg}`);
  }
  return parsed;
}

function printHelp() {
  console.log(`Render README GIFs from the real taskpane preview harness.

Usage:
  bun scripts/render-readme-gifs.mjs [options]

Options:
  --surface word,excel       Render a subset. Default: all six surfaces.
  --out <dir>                Output directory. Default: docs/assets/readme/sidepanes
  --port <port>              Preview server port. Default: ${DEFAULT_PORT}
  --check                    Verify expected GIF files exist and are non-empty.
  -h, --help                 Show this help.
`);
}

function startPreviewServer(portNumber) {
  const child = spawn(
    'bun',
    [
      'run',
      'vite',
      '--config',
      'vite.preview.config.ts',
      '--host',
      '127.0.0.1',
      '--port',
      String(portNumber),
      '--strictPort',
      '--open',
      'false',
    ],
    {
      cwd: path.join(ROOT, 'packages/web-shell'),
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env, BROWSER: 'none' },
    },
  );

  let logs = '';
  child.stdout.on('data', (chunk) => {
    logs += chunk.toString();
  });
  child.stderr.on('data', (chunk) => {
    logs += chunk.toString();
  });
  child.logs = () => logs;
  return child;
}

async function waitForPreview(url) {
  const deadline = Date.now() + 30_000;
  let lastError;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${url}/preview.html`);
      if (res.ok) return;
      lastError = new Error(`HTTP ${res.status}`);
    } catch (err) {
      lastError = err;
    }
    await sleep(250);
  }
  throw new Error(`Preview server did not become ready at ${url}: ${lastError?.message ?? 'timeout'}`);
}

async function renderSurfaceGif(browserInstance, url, surface, outputFile) {
  const page = await browserInstance.newPage({
    viewport: { width: 392, height: 792 },
    deviceScaleFactor: 1,
  });
  await page.goto(`${url}/preview.html?capture=1&surface=${surface}&scene=readme`, {
    waitUntil: 'networkidle',
  });
  const frame = page.locator('.preview-frame');
  await frame.waitFor({ state: 'visible', timeout: 10_000 });
  await page.waitForTimeout(400);

  const frames = [];
  await pushFrame(frames, frame, 950);

  await hoverIfPresent(page, '.surface-actions-trigger');
  await pushFrame(frames, frame, 900);

  await hoverIfPresent(page, '.surface-detail-trigger');
  await pushFrame(frames, frame, 900);

  await page.mouse.move(2, 2);
  await page.locator('.thread-region').evaluate((el) => {
    el.scrollTop = Math.min(420, el.scrollHeight);
  });
  await page.waitForTimeout(250);
  await pushFrame(frames, frame, 900);

  const runSteps = page.locator('.run-steps-toggle').first();
  if ((await runSteps.count()) > 0 && (await runSteps.isVisible())) {
    await runSteps.click({ timeout: 2_000 });
    await page.waitForTimeout(250);
    await pushFrame(frames, frame, 900);
  }

  const planEffect = page.locator('.plan-effect-head').first();
  if ((await planEffect.count()) > 0 && (await planEffect.isVisible())) {
    await planEffect.click({ timeout: 2_000 });
    await page.waitForTimeout(250);
    await pushFrame(frames, frame, 1200);
  }

  await page.close();
  await writeGif(frames, outputFile);
}

async function hoverIfPresent(page, selector) {
  const target = page.locator(selector).first();
  if ((await target.count()) === 0) return;
  if (!(await target.isVisible())) return;
  await target.hover({ timeout: 2_000 });
  await page.waitForTimeout(250);
}

async function pushFrame(frames, locator, delay) {
  const png = PNG.sync.read(await locator.screenshot({ type: 'png' }));
  frames.push({ width: png.width, height: png.height, data: png.data, delay });
}

async function writeGif(frames, outputFile) {
  if (frames.length === 0) throw new Error(`No frames captured for ${outputFile}`);
  const first = frames[0];
  const gif = GIFEncoder({ initialCapacity: first.width * first.height * frames.length });

  for (const frame of frames) {
    if (frame.width !== first.width || frame.height !== first.height) {
      throw new Error(`Frame size drifted while rendering ${outputFile}`);
    }
    const palette = quantize(frame.data, 128);
    const index = applyPalette(frame.data, palette);
    gif.writeFrame(index, frame.width, frame.height, {
      palette,
      delay: frame.delay,
      repeat: 0,
    });
  }

  gif.finish();
  await writeFile(outputFile, gif.bytes());
}

async function writeManifest(outDirPath, surfaces) {
  const lines = [
    '# README sidepane GIFs',
    '',
    'Generated from the real web-shell preview harness:',
    '',
    '```bash',
    'bun run docs:gifs',
    '```',
    '',
    '| Surface | Asset |',
    '| --- | --- |',
    ...surfaces.map((surface) => `| ${labelSurface(surface)} | \`${surface}.gif\` |`),
    '',
  ];
  await writeFile(path.join(outDirPath, 'README.md'), `${lines.join('\n')}\n`);
}

async function chromiumLaunchOptions() {
  const executablePath = await findChromiumExecutable();
  return executablePath ? { headless: true, executablePath } : { headless: true };
}

async function findChromiumExecutable() {
  const explicit = process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE;
  if (explicit && (await exists(explicit))) return explicit;

  const candidates = [
    process.env.HOME
      ? path.join(
          process.env.HOME,
          '.cache/ms-playwright/chromium_headless_shell-1228/chrome-headless-shell-linux64/chrome-headless-shell',
        )
      : '',
    process.env.HOME
      ? path.join(process.env.HOME, '.cache/ms-playwright/chromium-1228/chrome-linux64/chrome')
      : '',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (await exists(candidate)) return candidate;
  }

  const cacheRoot = process.env.HOME ? path.join(process.env.HOME, '.cache/ms-playwright') : '';
  if (!cacheRoot || !(await exists(cacheRoot))) return undefined;
  for (const entry of await readdir(cacheRoot, { withFileTypes: true })) {
    if (!entry.isDirectory() || !entry.name.startsWith('chromium-')) continue;
    const candidate = path.join(cacheRoot, entry.name, 'chrome-linux64/chrome');
    if (await exists(candidate)) return candidate;
  }
  return undefined;
}

async function checkExisting(outDirPath, surfaces) {
  for (const surface of surfaces) {
    const file = path.join(outDirPath, `${surface}.gif`);
    await access(file);
    const size = await stat(file);
    if (size.size <= 0) throw new Error(`${file} is empty`);
    console.log(`ok ${path.relative(ROOT, file)} (${Math.round(size.size / 1024)} KiB)`);
  }
}

async function exists(file) {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

function labelSurface(surface) {
  return {
    word: 'Word',
    excel: 'Excel',
    powerpoint: 'PowerPoint',
    outlook: 'Outlook',
    onenote: 'OneNote',
    teams: 'Teams',
  }[surface];
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function stopServer(child) {
  if (child.exitCode !== null || child.signalCode !== null) return;
  child.kill('SIGTERM');
  await Promise.race([
    new Promise((resolve) => child.once('exit', resolve)),
    sleep(3_000).then(() => {
      if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL');
    }),
  ]);
}
