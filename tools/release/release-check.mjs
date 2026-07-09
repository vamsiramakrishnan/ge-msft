#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import {
  artifactPlaceholderFindings,
  command,
  generatedManifestPath,
  gitDirty,
  gitSha,
  bunVersion,
  nodeVersion,
  packageDir,
  packageJson,
  packageZip,
  parseArgs,
  profileFromArgs,
  releaseConfig,
  repoRoot,
  rootVersion,
  scanForbiddenSecrets,
  sha256File,
  validateGeneratedManifest,
  verifyChecksums,
  walk,
  writeJson,
} from './common.mjs';

const args = parseArgs();
const profile = profileFromArgs(args);
const outDir = join(repoRoot, 'dist', 'release');

function ok(name, detail = '') {
  return { name, status: 'pass', detail };
}

function fail(name, detail) {
  return { name, status: 'fail', detail };
}

function blocked(name, detail, action) {
  return { name, status: 'blocked', detail, ...(action ? { action } : {}) };
}

function fromCommand(name, res) {
  return {
    name,
    status: res.ok ? 'pass' : 'fail',
    detail: `${res.command} exited ${res.status}`,
    durationMs: res.durationMs,
    stdoutTail: tail(res.stdout),
    stderrTail: tail(res.stderr),
  };
}

function tail(text) {
  return text.split(/\r?\n/).filter(Boolean).slice(-40).join('\n');
}

function runQualityChecks() {
  return [
    fromCommand('typecheck', command('bun', ['run', 'typecheck'])),
    fromCommand('lint', command('bun', ['run', 'lint'])),
    fromCommand('test', command('bun', ['run', 'test'])),
    fromCommand('coverage', command('bun', ['run', 'coverage'])),
    fromCommand('build', command('bun', ['run', 'build'])),
  ];
}

function runPythonChecks() {
  return [
    fromCommand(
      'python parse_commands self-test',
      command('python3', ['skill/m365-surface-commander/scripts/parse_commands.py', '--self-test']),
    ),
    fromCommand(
      'python surface_cli self-test',
      command('python3', ['skill/m365-surface-commander/scripts/surface_cli.py', '--self-test']),
    ),
    fromCommand(
      'python planner parity',
      command('python3', ['parity_test.py'], { cwd: join(repoRoot, 'skill') }),
    ),
    fromCommand(
      'python corpus parity',
      command('python3', ['parity_corpus_test.py'], { cwd: join(repoRoot, 'skill') }),
    ),
    fromCommand(
      'python tooling tests',
      command('python3', ['test_tooling.py'], { cwd: join(repoRoot, 'skill') }),
    ),
    fromCommand(
      'python deterministic eval',
      command('python3', ['eval/eval_harness.py'], { cwd: join(repoRoot, 'skill') }),
    ),
  ];
}

function manifestOutcome() {
  const manifestPath = generatedManifestPath(profile);
  if (!existsSync(manifestPath)) {
    const generated = command('bun', ['run', 'manifests:generate', '--', '--profile', profile]);
    if (!generated.ok) {
      const isExternal =
        generated.status === 2 ||
        /BLOCKED_EXTERNAL|Missing release manifest configuration/i.test(generated.stderr);
      return [
        isExternal
          ? blocked(
              'manifest generation',
              tail(generated.stderr || generated.stdout),
              'Set GE_ALPHA_APP_ID, GE_ALPHA_ENTRA_CLIENT_ID, GE_ALPHA_WEB_DOMAIN, GE_ALPHA_DEVELOPER_NAME, GE_ALPHA_WEBSITE_URL, GE_ALPHA_PRIVACY_URL, GE_ALPHA_TERMS_URL, and GE_ALPHA_SUPPORT_URL, then rerun bun run manifests:generate -- --profile internal-alpha-word-excel.',
            )
          : fail('manifest generation', tail(generated.stderr || generated.stdout)),
      ];
    }
  }
  const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
  const errors = validateGeneratedManifest(manifest, profile);
  return [
    errors.length
      ? fail('manifest validation', errors.join('; '))
      : ok('manifest validation', relative(repoRoot, manifestPath)),
  ];
}

function packageOutcome() {
  const zip = packageZip(profile);
  if (!existsSync(zip)) {
    const script = profile === 'development' ? 'package:dev' : 'package:alpha';
    const packaged = command('bun', ['run', script]);
    if (!packaged.ok) return [fail(`${profile} package`, tail(packaged.stderr || packaged.stdout))];
  }
  const checksumFailures = verifyChecksums(join(outDir, 'SHA256SUMS'));
  return [
    existsSync(zip)
      ? ok(`${profile} package`, `${relative(repoRoot, zip)} ${sha256File(zip)}`)
      : fail(`${profile} package`, 'package zip missing'),
    checksumFailures.length
      ? fail('artifact checksum verification', JSON.stringify(checksumFailures))
      : ok('artifact checksum verification', 'SHA256SUMS verified'),
  ];
}

function scanOutcomes() {
  const artifactRoots = [
    join(repoRoot, 'dist', 'manifests'),
    packageDir(profile),
    join(repoRoot, 'packages', 'web-shell', 'dist-web'),
  ].filter((p) => existsSync(p));
  const secretFindings = scanForbiddenSecrets(artifactRoots);
  const placeholders = artifactPlaceholderFindings(
    [join(repoRoot, 'dist', 'manifests'), packageDir(profile)].filter((p) => existsSync(p)),
  );
  return [
    secretFindings.length
      ? fail('browser bundle forbidden-secret scan', JSON.stringify(secretFindings))
      : ok('browser bundle forbidden-secret scan', `${artifactRoots.length} artifact root(s)`),
    placeholders.length
      ? fail('placeholder scan', placeholders.join(', '))
      : ok('placeholder scan', 'release artifacts contain no placeholders/dev hosts'),
  ];
}

function dependencyOutcomes() {
  const audit = command('bun', ['audit', '--audit-level=high', '--json']);
  const networkBlocked =
    /ENOTFOUND|ECONN|EAI_AGAIN|network|fetch failed/i.test(audit.stderr + audit.stdout) ||
    /ConnectionRefused|Connection refused|audit request failed/i.test(audit.stderr + audit.stdout);
  const outcomes = [
    networkBlocked
      ? blocked(
          'dependency/security scan',
          tail(audit.stderr || audit.stdout),
          'Run bun audit --audit-level=high --json from a network-enabled CI runner.',
        )
      : audit.ok
        ? ok('dependency/security scan', 'bun audit found no high+ vulnerabilities')
        : fail('dependency/security scan', tail(audit.stdout || audit.stderr)),
  ];
  const lockPath = join(repoRoot, 'bun.lock');
  if (!existsSync(lockPath)) {
    outcomes.push(fail('bun lockfile', 'bun.lock missing; run bun install'));
    return outcomes;
  }
  const lockText = readFileSync(lockPath, 'utf8');
  const packages = parseBunLockPackages(lockText);
  writeJson(join(outDir, 'license-inventory.json'), packages);
  writeJson(join(outDir, 'sbom.cdx.json'), {
    bomFormat: 'CycloneDX',
    specVersion: '1.5',
    version: 1,
    metadata: {
      component: { type: 'application', name: packageJson().name, version: rootVersion() },
    },
    components: packages.map((p) => ({ type: 'library', name: p.name, version: p.version })),
  });
  outcomes.push(ok('license inventory', 'dist/release/license-inventory.json'));
  outcomes.push(ok('SBOM', 'dist/release/sbom.cdx.json'));
  return outcomes;
}

function parseBunLockPackages(lockText) {
  const packages = new Map();
  const packageEntry = /^\s+"([^"]+)":\s+\["([^"]+)"/gm;
  for (const match of lockText.matchAll(packageEntry)) {
    const name = match[1];
    const version = match[2];
    packages.set(name, { name, version, license: 'UNKNOWN' });
  }
  return [...packages.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function authConfigOutcome() {
  const required = [
    'VITE_GCP_PROJECT',
    'VITE_GCP_LOCATION',
    'VITE_GE_ENGINE',
    'VITE_WIF_POOL_ID',
    'VITE_WIF_PROVIDER_ID',
    'VITE_ENTRA_TENANT_ID',
    'VITE_ENTRA_CLIENT_ID',
  ];
  const missing = required.filter((key) => !process.env[key]);
  if (missing.length) {
    return blocked(
      'authentication/WIF configuration',
      `missing ${missing.join(', ')}`,
      'Provide the production VITE_* public browser configuration in CI; do not provide client secrets.',
    );
  }
  return ok(
    'authentication/WIF configuration',
    'required public WIF/Entra coordinates are present',
  );
}

function certificationOutcomes(commitSha, packageHash, manifestVersion) {
  const evidenceDir = join(repoRoot, 'evidence', 'host-certification');
  const files = walk(evidenceDir).filter((f) => f.endsWith('.json'));
  const surfaces = ['word', 'excel'];
  const outcomes = [];
  for (const surface of surfaces) {
    const report = files
      .map((file) => {
        try {
          return { file, data: JSON.parse(readFileSync(file, 'utf8')) };
        } catch {
          return undefined;
        }
      })
      .filter(Boolean)
      .find(({ data }) => data.surface === surface && data.profile === profile);
    if (!report) {
      outcomes.push(
        blocked(
          `live-host certification ${surface}`,
          'no matching evidence report',
          `Run the ${surface} live-host certification protocol from docs/release/HOST-CERTIFICATION.md against this package and commit, then place the JSON report under evidence/host-certification/.`,
        ),
      );
      continue;
    }
    const { data, file } = report;
    const errors = [];
    if (data.commitSha !== commitSha) errors.push('commitSha mismatch');
    if (data.packageHash !== packageHash) errors.push('packageHash mismatch');
    if (data.manifestVersion !== manifestVersion) errors.push('manifestVersion mismatch');
    const failing = (data.tests ?? []).filter((t) => t.status !== 'pass');
    if (failing.length) errors.push(`${failing.length} test(s) not pass`);
    outcomes.push(
      errors.length
        ? fail(
            `live-host certification ${surface}`,
            `${relative(repoRoot, file)}: ${errors.join('; ')}`,
          )
        : ok(`live-host certification ${surface}`, relative(repoRoot, file)),
    );
  }
  return outcomes;
}

function provenanceOutcome() {
  const word = existsSync(join(repoRoot, 'packages', 'bridge-word', 'src', 'provenance-record.ts'));
  const excel = existsSync(
    join(repoRoot, 'packages', 'bridge-excel', 'src', 'provenance-record.ts'),
  );
  return word && excel
    ? ok(
        'durable provenance static paths',
        'Word custom XML and Excel workbook settings paths exist',
      )
    : fail('durable provenance static paths', 'Word or Excel provenance path missing');
}

function rollbackOutcome() {
  return existsSync(join(repoRoot, 'docs', 'release', 'ROLLBACK.md'))
    ? ok('rollback artifact availability', 'docs/release/ROLLBACK.md')
    : fail('rollback artifact availability', 'docs/release/ROLLBACK.md missing');
}

function finalResult(outcomes) {
  if (outcomes.some((o) => o.status === 'fail')) return 'FAIL';
  if (outcomes.some((o) => o.status === 'blocked')) return 'BLOCKED_EXTERNAL';
  return 'PASS';
}

function renderMarkdown(report) {
  const lines = [
    `# Release Status`,
    '',
    `Commit: ${report.commitSha}`,
    `Profile: ${report.profile}`,
    `Final result: ${report.finalResult}`,
    '',
    `| Gate | Status | Detail |`,
    `|---|---:|---|`,
  ];
  for (const outcome of report.outcomes) {
    lines.push(
      `| ${outcome.name} | ${outcome.status} | ${(outcome.detail ?? '').replace(/\n/g, '<br>')} |`,
    );
  }
  if (report.unresolvedBlockers.length) {
    lines.push('', '## Unresolved Blockers');
    for (const b of report.unresolvedBlockers) lines.push(`- ${b.name}: ${b.action ?? b.detail}`);
  }
  return `${lines.join('\n')}\n`;
}

const commitSha = gitSha();
const dirty = gitDirty();
const manifestPath = generatedManifestPath(profile);
const packagePath = packageZip(profile);
const outcomes = [
  dirty
    ? fail('clean working tree', 'working tree has uncommitted changes')
    : ok('clean working tree'),
];

outcomes.push(...runQualityChecks());
outcomes.push(...runPythonChecks());
outcomes.push(...manifestOutcome());
outcomes.push(...packageOutcome());
outcomes.push(...scanOutcomes());
outcomes.push(...dependencyOutcomes());
outcomes.push(authConfigOutcome());
outcomes.push(provenanceOutcome());
outcomes.push(rollbackOutcome());

let manifestVersion = 'unknown';
if (existsSync(manifestPath))
  manifestVersion = JSON.parse(readFileSync(manifestPath, 'utf8')).version;
const packageHash = existsSync(packagePath) ? sha256File(packagePath) : 'missing';
outcomes.push(...certificationOutcomes(commitSha, packageHash, manifestVersion));

const report = {
  schemaVersion: 1,
  commitSha,
  dirty,
  profile,
  packageVersion: rootVersion(),
  manifestVersion,
  nodeVersion: nodeVersion(),
  bunVersion: bunVersion(),
  enabledSurfaces:
    profile === 'internal-alpha-word-excel'
      ? ['word', 'excel']
      : ['word', 'excel', 'powerpoint', 'onenote', 'outlook', 'teams'],
  enabledReadCapabilities: {
    word: ['outline', 'read', 'search'],
    excel: ['outline', 'read', 'search'],
  },
  enabledWriteCapabilities: {
    word: ['tracked-change', 'add-comment', 'comment-reply'],
    excel: ['write-cells', 'format-cells', 'add-comment', 'comment-reply'],
  },
  artifacts: {
    manifest: existsSync(manifestPath) ? relative(repoRoot, manifestPath) : null,
    package: existsSync(packagePath) ? relative(repoRoot, packagePath) : null,
    packageSha256: packageHash,
  },
  outcomes,
  unresolvedBlockers: outcomes.filter((o) => o.status === 'blocked'),
  finalResult: finalResult(outcomes),
};

writeJson(join(outDir, 'release-status.json'), report);
writeFileSync(join(outDir, 'release-status.md'), renderMarkdown(report));
console.log(
  JSON.stringify(
    { finalResult: report.finalResult, report: 'dist/release/release-status.json' },
    null,
    2,
  ),
);
process.exit(report.finalResult === 'PASS' ? 0 : 1);
