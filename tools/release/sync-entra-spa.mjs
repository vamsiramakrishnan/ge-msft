#!/usr/bin/env node
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { spawnSync } from 'node:child_process';
import { parseArgs, releaseConfig, repoRoot } from './common.mjs';

const args = parseArgs();
const profile = String(args.profile ?? 'development');
const dryRun = Boolean(args['dry-run']);
const keepStaleRedirects = Boolean(args['keep-stale-redirects']);
const json = Boolean(args.json);

function command(cmd, commandArgs, options = {}) {
  const res = spawnSync(cmd, commandArgs, {
    cwd: repoRoot,
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe'],
    env: { ...process.env, ...(options.env ?? {}) },
  });
  return {
    ok: res.status === 0,
    status: res.status ?? 1,
    stdout: res.stdout ?? '',
    stderr: res.stderr ?? '',
    command: [cmd, ...commandArgs].join(' '),
  };
}

function fail(message, detail = '') {
  if (json) {
    console.log(JSON.stringify({ ok: false, message, detail }, null, 2));
  } else {
    console.error(message);
    if (detail) console.error(detail);
  }
  process.exit(1);
}

function findAz() {
  const candidates = [
    process.env.AZ_BIN,
    join(repoRoot, 'bin', 'az'),
    join(repoRoot, '.venv-az', 'bin', 'az'),
    'az',
  ].filter(Boolean);
  for (const candidate of candidates) {
    if (candidate.includes('/')) {
      if (existsSync(candidate)) return candidate;
      continue;
    }
    const found = command('bash', ['-lc', `command -v ${candidate}`]);
    if (found.ok && found.stdout.trim()) return found.stdout.trim();
  }
  return '';
}

function tail(text) {
  return String(text).split(/\r?\n/).filter(Boolean).slice(-12).join('\n');
}

function writeJson(path, value) {
  writeFileSync(path, `${JSON.stringify(value, null, 2)}\n`);
}

let cfg;
try {
  cfg = releaseConfig(profile);
} catch (error) {
  fail(error instanceof Error ? error.message : String(error));
}

const origin = String(args.origin ?? cfg.webOrigin ?? `https://${cfg.webDomain}`).replace(
  /\/+$/,
  '',
);
const redirect = String(args.redirect ?? `${origin}/auth-redirect.html`);
const clientId = String(args['client-id'] ?? cfg.entraClientId ?? '');

if (!clientId) fail('No Entra client id found for SPA redirect sync.');
if (
  !/^https:\/\/[^/]+\/auth-redirect\.html$/i.test(redirect) &&
  !/^http:\/\/localhost:\d+\/auth-redirect\.html$/i.test(redirect)
) {
  fail(`Refusing unexpected redirect URI: ${redirect}`);
}

const az = findAz();
if (!az) fail('Azure CLI not found. Install az or use the repo-local bin/az wrapper.');

const account = command(az, ['account', 'show']);
if (!account.ok) {
  fail(
    'Azure CLI is not signed in.',
    `Run: ${az} login --use-device-code\n${tail(account.stderr || account.stdout)}`,
  );
}

const encodedClientId = encodeURIComponent(clientId);
const queryUrl =
  `https://graph.microsoft.com/v1.0/applications?` +
  `$filter=appId%20eq%20%27${encodedClientId}%27&$select=id,appId,displayName,spa`;
const get = command(az, ['rest', '--only-show-errors', '--method', 'GET', '--url', queryUrl]);
if (!get.ok)
  fail('Failed to read Entra application through Microsoft Graph.', tail(get.stderr || get.stdout));

let app;
try {
  const parsed = JSON.parse(get.stdout);
  app = parsed.value?.[0];
} catch (error) {
  fail(
    'Failed to parse Microsoft Graph application response.',
    error instanceof Error ? error.message : String(error),
  );
}
if (!app) fail(`No Entra application matched appId ${clientId}.`);

const current = Array.isArray(app.spa?.redirectUris) ? app.spa.redirectUris : [];
const staleTrycloudflare = /^https:\/\/[^/]+\.trycloudflare\.com\/auth-redirect\.html$/i;
const kept = keepStaleRedirects ? current : current.filter((uri) => !staleTrycloudflare.test(uri));
const redirectUris = Array.from(new Set([...kept, redirect]));
const patch = { spa: { redirectUris } };
const patchPath = join(repoRoot, '.ge-dev', `entra-spa-patch.${profile}.json`);

mkdirSync(dirname(patchPath), { recursive: true });
if (!dryRun) {
  writeFileSync(join(dirname(patchPath), '.keep'), '');
}
writeJson(patchPath, patch);

const result = {
  ok: true,
  profile,
  appObjectId: app.id,
  appId: app.appId,
  displayName: app.displayName || '',
  redirect,
  before: current.length,
  after: redirectUris.length,
  removedStaleTrycloudflare: current.length - kept.length,
  patchPath: relative(repoRoot, patchPath),
  dryRun,
};

if (!dryRun) {
  const patchResult = command(az, [
    'rest',
    '--only-show-errors',
    '--method',
    'PATCH',
    '--url',
    `https://graph.microsoft.com/v1.0/applications/${app.id}`,
    '--headers',
    'Content-Type=application/json',
    '--body',
    JSON.stringify(patch),
  ]);
  if (!patchResult.ok) {
    fail(
      'Failed to patch Entra SPA redirect URIs.',
      tail(patchResult.stderr || patchResult.stdout),
    );
  }
}

if (json) {
  console.log(JSON.stringify(result, null, 2));
} else {
  console.log(`profile:      ${result.profile}`);
  console.log(`app:          ${result.displayName || result.appId}`);
  console.log(`redirect:     ${result.redirect}`);
  console.log(`redirects:    ${result.before} -> ${result.after}`);
  console.log(`removed:      ${result.removedStaleTrycloudflare} stale trycloudflare redirect(s)`);
  console.log(`patch file:   ${result.patchPath}`);
  if (dryRun) console.log('dry-run:      no Entra mutation performed');
}
