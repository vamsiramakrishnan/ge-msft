#!/usr/bin/env node
/**
 * Interactive preflight for the live StreamAssist contract suite.
 *
 * This does not automate Google cookies or XSRF state. It opens/prints the Gemini Enterprise app
 * URL, asks the developer to paste one authenticated DevTools cURL/HAR request, extracts only the
 * short-lived widget bearer/config via skill/extract_widget_credentials.py, then runs Vitest.
 */

import { existsSync, readFileSync } from 'node:fs';
import { createInterface } from 'node:readline/promises';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const defaultEnvFile = '/tmp/ge-widget.env';
const defaultTokenFile = '/tmp/ge-widget-token';

const args = parseArgs(process.argv.slice(2));
const envFile = args.envFile ?? process.env.GE_WIDGET_ENV_FILE ?? defaultEnvFile;
const tokenFile = args.tokenFile ?? process.env.GE_WIDGET_BEARER_TOKEN_FILE ?? defaultTokenFile;
const fileEnv = {
  ...readEnvFile(join(repoRoot, 'packages', 'web-shell', '.env')),
  ...readEnvFile(envFile),
};
const effectiveEnv = { ...fileEnv, ...process.env };

const token = tokenFromEnv(effectiveEnv, tokenFile);
const shouldRefresh = args.refresh || !isFreshWidgetToken(token);

if (shouldRefresh) {
  await refreshCredentials();
}

await runLiveTests();

function parseArgs(argv) {
  const parsed = {
    refresh: false,
    curlFile: undefined,
    envFile: undefined,
    tokenFile: undefined,
    scenarioCsv: undefined,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    switch (arg) {
      case '--refresh':
        parsed.refresh = true;
        break;
      case '--curl-file':
        parsed.curlFile = requireValue(argv, ++i, arg);
        break;
      case '--env-file':
        parsed.envFile = requireValue(argv, ++i, arg);
        break;
      case '--token-file':
        parsed.tokenFile = requireValue(argv, ++i, arg);
        break;
      case '--scenarios':
        parsed.scenarioCsv = requireValue(argv, ++i, arg);
        break;
      case '--help':
      case '-h':
        usage();
        process.exit(0);
        break;
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }
  return parsed;
}

function usage() {
  console.log(`streamassist-live-login

Refresh a short-lived Gemini Enterprise widget token, then run the live StreamAssist tests.

Usage:
  bun run test:streamassist:live:login
  bun scripts/streamassist-live-login.mjs --refresh
  bun scripts/streamassist-live-login.mjs --curl-file /tmp/widget-request.curl
  bun scripts/streamassist-live-login.mjs --scenarios smoke-basic,commander-excel-visualize

Options:
  --refresh              Force credential refresh even if the token file looks fresh.
  --curl-file <path>     Read a copied DevTools cURL/HAR request from a file.
  --env-file <path>      Write/read widget env exports here. Default: /tmp/ge-widget.env.
  --token-file <path>    Write/read widget bearer here. Default: /tmp/ge-widget-token.
  --scenarios <csv>      Set GE_LIVE_STREAMASSIST_SCENARIOS for the run.
`);
}

function requireValue(argv, index, flag) {
  const value = argv[index];
  if (!value) throw new Error(`${flag} requires a value`);
  return value;
}

function readEnvFile(path) {
  if (!path || !existsSync(path)) return {};
  const out = {};
  for (const line of readFileSync(path, 'utf8').split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const match = /^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)=(.*)$/.exec(trimmed);
    if (!match) continue;
    const key = match[1];
    const raw = match[2];
    if (!key || raw === undefined) continue;
    out[key] = unquote(raw.trim());
  }
  return out;
}

function unquote(value) {
  if (
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"))
  ) {
    return value.slice(1, -1).replace(/'"'"'/g, "'");
  }
  return value;
}

function tokenFromEnv(env, fallbackFile) {
  const inline = env.GE_WIDGET_BEARER_TOKEN || env.GE_TOKEN || env.GE_ACCESS_TOKEN;
  if (inline) return inline.trim();
  const file = env.GE_WIDGET_BEARER_TOKEN_FILE || fallbackFile;
  if (!file || !existsSync(file)) return '';
  return readFileSync(file, 'utf8').trim();
}

function isFreshWidgetToken(rawToken) {
  const payload = jwtPayload(rawToken);
  if (!payload?.exp) return false;
  const expiresAtMs = Number(payload.exp) * 1000;
  const skewMs = 90_000;
  return Number.isFinite(expiresAtMs) && expiresAtMs - Date.now() > skewMs;
}

function jwtPayload(rawToken) {
  const [, payload] = rawToken.split('.');
  if (!payload) return undefined;
  try {
    const normalized = payload.replace(/-/g, '+').replace(/_/g, '/');
    const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, '=');
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8'));
  } catch {
    return undefined;
  }
}

async function refreshCredentials() {
  const configId =
    process.env.GE_WIDGET_CONFIG_ID ||
    fileEnv.GE_WIDGET_CONFIG_ID ||
    process.env.VITE_GE_WIDGET_CONFIG_ID ||
    fileEnv.VITE_GE_WIDGET_CONFIG_ID;
  const loginUrl =
    process.env.GE_WIDGET_LOGIN_URL ||
    (configId ? `https://vertexaisearch.cloud.google/home/cid/${configId}` : 'https://vertexaisearch.cloud.google/');

  console.log('Refreshing Gemini Enterprise widget credentials.');
  console.log(`Open/sign in here:\n${loginUrl}\n`);
  tryOpenBrowser(loginUrl);

  const capturedRequest = args.curlFile ? readFileSync(args.curlFile, 'utf8') : await promptForCurl();
  if (!capturedRequest.trim()) {
    throw new Error('No DevTools cURL/HAR request provided; cannot refresh widget credentials.');
  }

  const py = spawnSync(
    'python3',
    [
      join(repoRoot, 'skill', 'extract_widget_credentials.py'),
      '-',
      '--env-file',
      envFile,
      '--token-file',
      tokenFile,
      ...(effectiveEnv.GE_PROJECT ? ['--project', effectiveEnv.GE_PROJECT] : []),
    ],
    {
      cwd: repoRoot,
      input: capturedRequest,
      encoding: 'utf8',
      stdio: ['pipe', 'inherit', 'inherit'],
    },
  );
  if (py.status !== 0) {
    throw new Error(`credential extraction failed with exit code ${py.status ?? 'unknown'}`);
  }
}

function tryOpenBrowser(url) {
  if (process.env.GE_STREAMASSIST_OPEN_BROWSER === '0') return;
  const result = spawnSync('python3', ['-m', 'webbrowser', url], {
    stdio: 'ignore',
  });
  if (result.status !== 0) {
    console.log('Browser auto-open failed or is unavailable; use the URL above.');
  }
}

async function promptForCurl() {
  if (!process.stdin.isTTY) {
    throw new Error(
      'No TTY available. Pass --curl-file /path/to/widget-request.curl for non-interactive runs.',
    );
  }
  console.log(
    [
      'After login:',
      '1. Open DevTools -> Network.',
      '2. Run any Gemini Enterprise request, e.g. list skills or ask a trivial question.',
      '3. Copy a content-discoveryengine.googleapis.com request as cURL, or export HAR text.',
      '4. Paste it below, then press Ctrl-D.',
      '',
    ].join('\n'),
  );
  const rl = createInterface({ input: process.stdin, output: process.stdout, terminal: false });
  const lines = [];
  for await (const line of rl) lines.push(line);
  return lines.join('\n');
}

async function runLiveTests() {
  const childEnv = {
    ...process.env,
    GE_LIVE_STREAMASSIST: '1',
    GE_WIDGET_ENV_FILE: envFile,
    GE_WIDGET_BEARER_TOKEN_FILE: tokenFile,
    ...(args.scenarioCsv ? { GE_LIVE_STREAMASSIST_SCENARIOS: args.scenarioCsv } : {}),
  };
  const child = spawn(
    process.platform === 'win32' ? 'bun.exe' : 'bun',
    [
      'run',
      'vitest',
      'run',
      'packages/gemini-client/src/stream-assist.live.test.ts',
      '--testTimeout=120000',
    ],
    {
      cwd: repoRoot,
      stdio: 'inherit',
      env: childEnv,
    },
  );
  const code = await new Promise((resolve) => child.on('exit', resolve));
  process.exit(typeof code === 'number' ? code : 1);
}
